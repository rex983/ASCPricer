import type {
  BuildingConfig,
  StandardMatrices,
  WidespanMatrices,
  PricingLookup,
} from "@/types/pricing";
import { lookupMatrix, lookupValue, nearestBucket } from "./lookups";
import {
  HEIGHT_MULTIPLIERS,
  WIND_LOAD_CATEGORIES,
  WIDESPAN_WIND_CATEGORIES_MAIN,
  WIDESPAN_WIND_CATEGORIES_GIRT,
} from "./constants";

// ── Helpers ──

/** Standard truss spacing buckets used in hat channel / girt lookups */
const TRUSS_SPACING_BUCKETS = [36, 42, 48, 54, 60] as const;

/** Get the height multiplier for snow engineering costs. */
export function getHeightMultiplier(height: number): number {
  if (height >= 19) return HEIGHT_MULTIPLIERS["19-20"];
  if (height >= 16) return HEIGHT_MULTIPLIERS["16-18"];
  if (height >= 13) return HEIGHT_MULTIPLIERS["13-15"];
  return HEIGHT_MULTIPLIERS["6-12"];
}

/** Resolve height to S/M/T prefix using heightClassification lookup */
function getHeightPrefix(
  height: number,
  classification: PricingLookup
): string {
  // Try exact height match
  const val = classification[String(height)];
  if (val !== undefined) {
    if (val === 0) return "S";
    if (val === 1) return "M";
    if (val === 2) return "T";
  }
  // Fallback: height-based classification
  if (height <= 12) return "S";
  if (height <= 15) return "M";
  return "T";
}

/** Bucket wind MPH to nearest standard category */
function bucketWind(
  inputMph: number,
  buckets: PricingLookup,
  categories: readonly number[]
): number {
  // Try exact lookup first
  const bucketed = buckets[String(inputMph)];
  if (bucketed && bucketed > 0) return bucketed;
  // Fallback: nearest bucket from categories
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
  // Default
  if (height <= 9) return 3;
  if (height <= 12) return 4;
  if (height <= 15) return 4;
  return 5;
}

/** Look up truss price for a given width and state from the per-state pricing matrix */
function getTrussPrice(
  width: number,
  state: string,
  trussPriceByWidthState: Record<string, Record<string, number>>
): number {
  const stateRow = trussPriceByWidthState[state];
  if (!stateRow) return 190; // fallback

  // Check each width range key (e.g. "12-24", "26-30")
  for (const [rangeKey, price] of Object.entries(stateRow)) {
    const parts = rangeKey.split("-").map(Number);
    if (parts.length === 2 && width >= parts[0] && width <= parts[1]) {
      return price;
    }
    // Exact width match
    if (parts.length === 1 && parts[0] === width) {
      return price;
    }
  }
  return 190; // fallback
}

/** Count enclosed vertical surfaces (sides only — girts go on sides) */
function countEnclosedVerticalSurfaces(config: BuildingConfig): number {
  let count = 0;
  if (config.sidesCoverage !== "open") count += config.sidesQty;
  return count;
}

/** Calculate vertical-surface perimeter length (sides only, for girts) */
function getVerticalPerimeter(config: BuildingConfig, length: number): number {
  const enclosedSides = countEnclosedVerticalSurfaces(config);
  return enclosedSides * length;
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
 * Step 1: Resolve inputs (bucket wind, height prefix, snow code, config key)
 * Step 2: Extra trusses
 * Step 3: Extra hat channels (two-stage lookup)
 * Step 4: Extra girts (vertical sides only)
 * Step 5: Extra verticals (uses WIDTH not height)
 * Step 6: Apply height multiplier to total
 */
export function calculateStandardSnowEngineering(
  config: BuildingConfig,
  matrices: StandardMatrices,
  resolvedKeys: { width: number; length: number; roofKey: string }
): number {
  if (!config.snowLoad) return 0;

  const { width, length, roofKey } = resolvedKeys;
  const state = config.state || "";

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
      const pricePerTruss = getTrussPrice(
        width,
        state,
        matrices.snow.trussPriceByWidthState
      );
      totalCost += extraTrusses * pricePerTruss;
    }
  }

  // ── Step 3: Extra Hat Channels (two-stage lookup) ──
  // Bucket truss spacing to nearest of (36/42/48/54/60)
  const actualTrussSpacing = trussSpacing > 0 ? trussSpacing : 60;
  const bucketedTrussSpacing = nearestBucket(
    actualTrussSpacing,
    TRUSS_SPACING_BUCKETS
  );

  // Row key: "{bucketedTrussSpacing}-{snowCode}"
  const hcRowKey = `${bucketedTrussSpacing}-${snowCode}`;
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

    // Original HC count — from hatChannelCounts matrix
    const widthStateKey = `${width}-${state}`;
    let originalChannels = lookupMatrix(
      matrices.snow.hatChannelCounts,
      widthStateKey,
      "count"
    );
    // Fallback: try just width
    if (originalChannels === 0) {
      originalChannels = lookupMatrix(
        matrices.snow.hatChannelCounts,
        String(width),
        "count"
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

  // ── Step 4: Extra Girts (only vertical/enclosed sides) ──
  const girtSpacing = lookupMatrix(
    matrices.snow.girtSpacing,
    String(bucketedTrussSpacing),
    String(bucketedWind)
  );
  if (girtSpacing > 0) {
    const heightInches = config.height * 12;
    const girtsNeeded = Math.ceil(heightInches / girtSpacing) + 1;
    const originalGirts = resolveOriginalGirts(
      config.height,
      matrices.snow.girtCountsByHeight
    );

    const extraGirts = Math.max(0, girtsNeeded - originalGirts);
    if (extraGirts > 0) {
      const tubingPrice =
        lookupValue(matrices.snow.tubingPriceByState, state) || 3;
      // Girts only on vertical (side) surfaces
      const perimeter = getVerticalPerimeter(config, length);
      totalCost += extraGirts * tubingPrice * perimeter;
    }
  }

  // ── Step 5: Extra Verticals (uses WIDTH not height) ──
  const verticalSpacing = lookupMatrix(
    matrices.snow.verticalSpacing,
    String(bucketedWind),
    String(config.height)
  );
  if (verticalSpacing > 0) {
    const widthInches = width * 12;
    const verticalsNeeded = Math.ceil(widthInches / verticalSpacing) + 1;

    const originalVerticals = lookupMatrix(
      matrices.snow.verticalCounts,
      String(width),
      "count"
    );

    const extraVerticals = Math.max(0, verticalsNeeded - originalVerticals);
    if (extraVerticals > 0) {
      const tubingPrice =
        lookupValue(matrices.snow.tubingPriceByState, state) || 3;
      // Verticals run the peak height
      const peakHeight = config.height; // ft
      totalCost += extraVerticals * tubingPrice * peakHeight;
    }
  }

  // ── Step 6: Apply height multiplier to total ──
  const heightMult = getHeightMultiplier(config.height);
  return Math.round(totalCost * heightMult);
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
