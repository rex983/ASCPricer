import type {
  BuildingConfig,
  StandardMatrices,
  WidespanMatrices,
  PricingLookup,
} from "@/types/pricing";
import { lookupMatrix, lookupValue, nearestBucket } from "./lookups";
import {
  WIND_LOAD_CATEGORIES,
  WIDESPAN_WIND_CATEGORIES_MAIN,
  WIDESPAN_WIND_CATEGORIES_GIRT,
} from "./constants";

// ── Helpers ──

/** Standard truss spacing buckets used in hat channel / girt lookups */
const TRUSS_SPACING_BUCKETS = [36, 42, 48, 54, 60] as const;

/**
 * Resolve a state code to one that exists in the truss counts matrix.
 * Each spreadsheet covers a region (e.g., AZ/CO/UT) but only has columns
 * for representative states (e.g., "24-AZ" but not "24-CO" or "24-UT").
 * If the given state isn't found, find the first available state for that width.
 */
function resolveEngineeringState(
  state: string,
  width: number,
  trussCounts: Record<string, Record<string, number>>
): string {
  // Try exact match first
  const exactKey = `${width}-${state}`;
  if (trussCounts[exactKey]) return state;

  // Find any state that exists for this width
  const prefix = `${width}-`;
  for (const key of Object.keys(trussCounts)) {
    if (key.startsWith(prefix)) {
      return key.slice(prefix.length);
    }
  }
  return state; // fallback to original
}

/** A-frame roof pitch: 3:12 (3 inches rise per 12 inches run) */
const ROOF_PITCH = 3 / 12;

/** Resolve height to S/M/T prefix using heightClassification lookup */
function getHeightPrefix(
  height: number,
  classification: PricingLookup
): string {
  const val = classification[String(height)];
  if (val !== undefined) {
    if (val === 0) return "S";
    if (val === 1) return "M";
    if (val === 2) return "T";
  }
  // Fallback
  if (height <= 6) return "S";
  if (height <= 9) return "M";
  return "T";
}

/** Bucket wind MPH to nearest standard category */
function bucketWind(
  inputMph: number,
  buckets: PricingLookup,
  categories: readonly number[]
): number {
  const bucketed = buckets[String(inputMph)];
  if (bucketed && bucketed > 0) return bucketed;
  return nearestBucket(inputMph, categories);
}

/** Resolve original girt count from girtCountsByHeight lookup */
function resolveOriginalGirts(
  height: number,
  girtCountsByHeight: PricingLookup
): number {
  // Try exact height
  const exact = girtCountsByHeight[String(height)];
  if (exact > 0) return exact;
  // Try range-based keys like "0-9", "10-12", "13-15", "16-20"
  for (const [key, count] of Object.entries(girtCountsByHeight)) {
    const parts = key.split("-").map(Number);
    if (parts.length === 2 && height >= parts[0] && height <= parts[1]) {
      return count;
    }
  }
  // Default (matches spreadsheet: 0-11→3, 12-17→4, 18+→5)
  if (height <= 11) return 3;
  if (height <= 17) return 4;
  return 5;
}

/** Look up truss price for a given width and state */
function getTrussPrice(
  width: number,
  state: string,
  trussPriceByWidthState: Record<string, Record<string, number>>
): number {
  const stateRow = trussPriceByWidthState[state];
  if (!stateRow) return 190; // fallback

  // Try exact width match first
  if (stateRow[String(width)] !== undefined) {
    return stateRow[String(width)];
  }

  // Try width range keys (e.g. "12-24", "26-30")
  for (const [rangeKey, price] of Object.entries(stateRow)) {
    const parts = rangeKey.split("-").map(Number);
    if (parts.length === 2 && width >= parts[0] && width <= parts[1]) {
      return price;
    }
  }
  return 190; // fallback
}

/** Calculate roof rise for A-frame roof styles */
function getRoofRise(width: number, roofKey: string): number {
  if (roofKey === "AFV" || roofKey === "AFH") {
    return (width / 2) * ROOF_PITCH; // e.g., 24W → 12 run × 0.25 = 3ft
  }
  return 0; // standard roof — flat/no rise for vertical calc purposes
}

// ── Standard Snow Engineering ──

/**
 * Build the snow engineering config key.
 * Format: "E-105-24-AFV" (enclosed/open - wind - width - roof)
 */
export function buildSnowConfigKey(
  isEnclosed: boolean,
  windLoad: number,
  width: number,
  roofKey: string
): string {
  const enclosure = isEnclosed ? "E" : "O";
  return `${enclosure}-${windLoad}-${width}-${roofKey}`;
}

/**
 * Calculate snow/wind engineering costs for standard buildings.
 *
 * Based on "Snow - Math Calculations" spreadsheet:
 * Step 1: Resolve inputs (bucket wind, height prefix, snow code, config key)
 * Step 2: Extra trusses (with leg surcharge via feetUsed × pieTrussPrice)
 * Step 3: Extra hat channels (two-stage lookup, NO height prefix in HC row key)
 * Step 4: Extra girts (ONLY if enclosed AND vertical panels)
 * Step 5: Extra verticals (uses WIDTH; pricing uses peakHeight = eave + roofRise)
 *
 * NO global height multiplier — each component handles height individually.
 */
export function calculateStandardSnowEngineering(
  config: BuildingConfig,
  matrices: StandardMatrices,
  resolvedKeys: { width: number; length: number; roofKey: string }
): number {
  if (!config.snowLoad) return 0;

  const { width, length, roofKey } = resolvedKeys;
  const rawState = config.state || "";

  // Resolve state to one that exists in the spreadsheet data
  // (e.g., CO/UT → AZ for the AZ/CO/UT region spreadsheet)
  const state = resolveEngineeringState(
    rawState,
    width,
    matrices.snow.trussCounts
  );

  // ── Step 1: Resolve inputs ──
  const bucketedWind = bucketWind(
    config.windRating || 105,
    matrices.snow.windLoadBuckets,
    WIND_LOAD_CATEGORIES
  );
  const heightPrefix = getHeightPrefix(
    config.height,
    matrices.snow.heightClassification
  );

  // Snow code: prepend height prefix → e.g., "T-20LL"
  const snowCode = `${heightPrefix}-${config.snowLoad}`;

  // Determine if building is enclosed
  const isEnclosed =
    config.sidesCoverage === "fully_enclosed" && config.endsQty >= 2;

  // Config key: "{E|O}-{bucketedWind}-{width}-{roofKey}"
  const configKey = buildSnowConfigKey(isEnclosed, bucketedWind, width, roofKey);

  let totalCost = 0;

  // ── Step 2: Extra Trusses ──
  const trussSpacing = lookupMatrix(
    matrices.snow.trussSpacing,
    snowCode,
    configKey
  );

  // If truss spacing = 0 for a valid snow load, the load exceeds standard
  // engineering → "Contact Engineer" (return -1 as sentinel)
  if (trussSpacing === 0) return -1;

  if (trussSpacing > 0) {
    const lengthInches = length * 12;
    const trussesNeeded = Math.ceil(lengthInches / trussSpacing) + 1;

    // Original truss count: trussCounts["{width}-{state}"]["{length}"]
    const widthStateKey = `${width}-${state}`;
    const originalTrusses = lookupMatrix(
      matrices.snow.trussCounts,
      widthStateKey,
      String(length)
    );

    const extraTrusses = Math.max(0, trussesNeeded - originalTrusses);
    if (extraTrusses > 0) {
      const baseTrussPrice = getTrussPrice(
        width,
        state,
        matrices.snow.trussPriceByWidthState
      );
      // Leg surcharge: feetUsed × pieTrussPrice per extra truss
      const feetUsed = lookupValue(
        matrices.snow.feetUsedByHeight,
        String(config.height)
      );
      const piePricePerFt = lookupValue(
        matrices.snow.pieTrussPrice,
        state
      ) || 15; // fallback $15
      const legSurcharge = feetUsed * piePricePerFt;
      totalCost += extraTrusses * (baseTrussPrice + legSurcharge);
    }
  }

  // ── Step 3: Extra Hat Channels (two-stage lookup) ──
  const actualTrussSpacing = trussSpacing > 0 ? trussSpacing : 60;
  const bucketedTrussSpacing = nearestBucket(
    actualTrussSpacing,
    TRUSS_SPACING_BUCKETS
  );

  // HC row key: "{bucketedTrussSpacing}-{snowLoad}" — NO height prefix!
  // e.g., "60-20LL" not "60-T-20LL"
  const hcRowKey = `${bucketedTrussSpacing}-${config.snowLoad}`;
  const hatChannelSpacing = lookupMatrix(
    matrices.snow.hatChannelSpacing,
    hcRowKey,
    String(bucketedWind)
  );

  if (hatChannelSpacing > 0) {
    const barSize = (width + 2) / 2; // half-width in ft
    const barInches = barSize * 12;
    const channelsPerSide = Math.ceil(barInches / hatChannelSpacing) + 1;
    const totalChannels = channelsPerSide * 2; // both sides of roof

    // Original HC count: hatChannelCounts[state][width]
    let originalChannels = lookupMatrix(
      matrices.snow.hatChannelCounts,
      state,
      String(width)
    );
    // Fallback: try width as row key
    if (originalChannels === 0) {
      originalChannels = lookupMatrix(
        matrices.snow.hatChannelCounts,
        String(width),
        state
      );
    }

    const extraChannels = Math.max(0, totalChannels - originalChannels);
    if (extraChannels > 0) {
      const channelPricePerFt =
        lookupValue(matrices.snow.channelPriceByState, state) || 2;
      const channelLength = length + 1;
      totalCost += extraChannels * channelPricePerFt * channelLength;
    }
  }

  // ── Step 4: Extra Girts ──
  // CRITICAL: Girts are ONLY needed if building is enclosed AND sides are vertical panels
  const girtsNeeded =
    isEnclosed && config.sidesOrientation === "vertical";

  if (girtsNeeded) {
    const girtSpacing = lookupMatrix(
      matrices.snow.girtSpacing,
      String(bucketedTrussSpacing),
      String(bucketedWind)
    );
    if (girtSpacing > 0) {
      const heightInches = config.height * 12;
      const girtsRequired = Math.ceil(heightInches / girtSpacing) + 1;
      const originalGirts = resolveOriginalGirts(
        config.height,
        matrices.snow.girtCountsByHeight
      );

      const extraGirts = Math.max(0, girtsRequired - originalGirts);
      if (extraGirts > 0) {
        const tubingPrice =
          lookupValue(matrices.snow.tubingPriceByState, state) || 3;
        // Girt perimeter: vertical sides × length + vertical ends × width
        let perimeter = 0;
        if (config.sidesCoverage !== "open" && config.sidesOrientation === "vertical") {
          perimeter += config.sidesQty * length;
        }
        if (config.endsQty > 0 && config.endsOrientation === "vertical") {
          perimeter += config.endsQty * width;
        }
        totalCost += extraGirts * tubingPrice * perimeter;
      }
    }
  }

  // ── Step 5: Extra Verticals ──
  // Vertical spacing: matrix[height][windSpeed] (readMatrix without transpose)
  const verticalSpacing = lookupMatrix(
    matrices.snow.verticalSpacing,
    String(config.height),
    String(bucketedWind)
  );
  if (verticalSpacing > 0) {
    // Verticals use WIDTH (not height) for count calculation
    const widthInches = width * 12;
    const verticalsNeeded = Math.ceil(widthInches / verticalSpacing) + 1;

    // Original vertical count: verticalCounts[width] (PricingLookup)
    const originalVerticals = lookupValue(
      matrices.snow.verticalCounts,
      String(width)
    );

    const extraVerticals = Math.max(0, verticalsNeeded - originalVerticals);
    if (extraVerticals > 0 && config.endsQty > 0) {
      const tubingPrice =
        lookupValue(matrices.snow.tubingPriceByState, state) || 3;
      // Verticals run the PEAK height (eave + roof rise)
      const peakHeight = config.height + getRoofRise(width, roofKey);
      // Verticals are installed at each enclosed end wall
      totalCost += extraVerticals * config.endsQty * tubingPrice * peakHeight;
    }
  }

  // NO height multiplier — each component handles height individually
  return Math.round(totalCost);
}

// ── Widespan Snow Engineering ──

/**
 * Calculate snow/wind engineering costs for widespan buildings.
 *
 * Key differences from standard:
 * 1. Two wind mappings: mapping #1 (105/120/130) for trusses+purlins, mapping #2 (90/110/120/130) for girts
 * 2. Has verticals (based on wind, not snow)
 * 3. Uses per-ft costs from snowCalc
 * 4. Girt perimeter based on enclosed surfaces only
 */
export function calculateWidespanSnowEngineering(
  config: BuildingConfig,
  matrices: WidespanMatrices,
  resolvedKeys: { width: number; length: number }
): number {
  if (!config.snowLoad) return 0;

  const { width, length } = resolvedKeys;

  // Determine enclosure and building type
  const isEnclosed =
    config.sidesCoverage === "fully_enclosed" && config.endsQty >= 2;
  const enclosure = isEnclosed ? "E" : "O";
  const buildingType = width >= 42 ? "G" : "S";

  // Wind load — mapping #1 for trusses/purlins (105/120/130)
  const rawWind = config.windRating || 105;
  const windMain = nearestBucket(rawWind, WIDESPAN_WIND_CATEGORIES_MAIN);
  // Wind load — mapping #2 for girts (90/110/120/130)
  const windGirt = nearestBucket(rawWind, WIDESPAN_WIND_CATEGORIES_GIRT);

  // Build config key: "E-105-S" or "O-130-G"
  const configKey = `${enclosure}-${windMain}-${buildingType}`;

  // Snow load code (e.g., "S-30GL", "G-50GL")
  const snowCode = `${buildingType}-${config.snowLoad}`;

  let totalCost = 0;

  // ── Extra Trusses ──
  const trussData = matrices.snow.trussCounts[String(length)];
  if (trussData) {
    const originalTrusses = trussData.count || 0;
    const maxSpacing =
      matrices.snow.trussSpacing[String(length)]?.spacing || 120;

    const lengthInches = length * 12;
    const trussesNeeded = Math.ceil(lengthInches / maxSpacing) + 1;
    const extraTrusses = Math.max(0, trussesNeeded - originalTrusses);

    if (extraTrusses > 0) {
      const trussPrice =
        lookupValue(matrices.snow.trussPriceByState, String(width)) || 2000;
      totalCost += extraTrusses * trussPrice;

      // Leg height cost for extra trusses
      if (config.height > 10) {
        const legCostPerFt = matrices.snow.legTrussCostPerFt || 90;
        totalCost += extraTrusses * (config.height - 10) * legCostPerFt;
      }
    }
  }

  // ── Extra Purlins ──
  const purlinSpacing = lookupMatrix(
    matrices.snow.purlinSpacing,
    configKey,
    snowCode
  );
  if (purlinSpacing > 0) {
    const halfWidth = (width + 2) / 2;
    const halfInches = halfWidth * 12;
    const purlinsNeeded = Math.ceil(halfInches / purlinSpacing) + 1;
    const totalPurlins = purlinsNeeded * 2;

    const originalPurlins =
      matrices.snow.purlinCounts[String(width)]?.count || 12;
    const extraPurlins = Math.max(0, totalPurlins - originalPurlins);

    if (extraPurlins > 0) {
      const purlinCostPerFt = matrices.snow.purlinCostPerFt || 6;
      totalCost += extraPurlins * purlinCostPerFt * length;
    }
  }

  // ── Extra Girts (uses wind mapping #2) ──
  const girtSpacingVal =
    lookupValue(matrices.snow.windLoadMapping2, String(windGirt)) ||
    matrices.snow.girtSpacing[String(windGirt)]?.spacing;
  if (girtSpacingVal && girtSpacingVal > 0) {
    const heightInches = config.height * 12;
    const girtsNeeded = Math.ceil(heightInches / girtSpacingVal) + 1;
    const originalGirts =
      matrices.snow.girtCounts[String(config.height)]?.count || 3;
    const extraGirts = Math.max(0, girtsNeeded - originalGirts);

    if (extraGirts > 0) {
      // Perimeter based on enclosed surfaces only
      let perimeter = 0;
      if (config.sidesCoverage !== "open") perimeter += config.sidesQty * length;
      if (config.endsQty > 0) perimeter += config.endsQty * width;
      const girtCostPerFt = matrices.snow.girtCostPerFt || 6;
      totalCost += extraGirts * girtCostPerFt * perimeter;
    }
  }

  // ── Extra Verticals ──
  const verticalSpacing = lookupValue(
    matrices.snow.verticalSpacingByWind,
    String(windMain)
  );
  if (verticalSpacing > 0) {
    const originalVerticals =
      lookupValue(matrices.snow.verticalCountByWidth, String(width)) || 0;

    const halfWidth = width / 2;
    const halfInches = halfWidth * 12;
    const verticalsNeeded = Math.ceil(halfInches / verticalSpacing) + 1;
    const totalVerticals = verticalsNeeded * 2; // both sides

    const extraVerticals = Math.max(0, totalVerticals - originalVerticals);
    if (extraVerticals > 0) {
      const verticalCostPerFt = matrices.snow.verticalCostPerFt || 18;
      totalCost += extraVerticals * verticalCostPerFt * halfWidth;
    }
  }

  return Math.round(totalCost);
}
