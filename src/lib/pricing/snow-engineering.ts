import type {
  BuildingConfig,
  StandardMatrices,
  WidespanMatrices,
} from "@/types/pricing";
import { lookupMatrix, lookupValue } from "./lookups";
import { HEIGHT_MULTIPLIERS } from "./constants";

/**
 * Get the height multiplier for snow engineering costs.
 */
export function getHeightMultiplier(height: number): number {
  if (height >= 19) return HEIGHT_MULTIPLIERS["19-20"];
  if (height >= 16) return HEIGHT_MULTIPLIERS["16-18"];
  if (height >= 13) return HEIGHT_MULTIPLIERS["13-15"];
  return HEIGHT_MULTIPLIERS["6-12"];
}

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
 * The calculation:
 * 1. Build config key → look up required truss spacing
 * 2. Calculate extra trusses: (length_inches / required_spacing) + 1 - original_count
 * 3. Same for hat channels, girts, verticals
 * 4. Price: extra_units × per_unit_cost × height_multiplier
 */
export function calculateStandardSnowEngineering(
  config: BuildingConfig,
  matrices: StandardMatrices,
  resolvedKeys: { width: number; length: number; roofKey: string }
): number {
  if (!config.snowLoad) return 0;

  const { width, length, roofKey } = resolvedKeys;
  const heightMult = getHeightMultiplier(config.height);

  // Determine if building is enclosed
  const isEnclosed =
    config.sidesCoverage === "fully_enclosed" && config.endsQty >= 2;

  // Build config key: "E-105-24-AFV"
  const windLoad = config.windRating || 105;
  const configKey = buildSnowConfigKey(isEnclosed, windLoad, width, roofKey);

  // Snow load code (e.g., "T-30GL", "T-20LL")
  const snowCode = config.snowLoad;

  let totalCost = 0;

  // ── Extra Trusses ──
  const trussSpacing = lookupMatrix(
    matrices.snow.trussSpacing,
    configKey,
    snowCode
  );
  if (trussSpacing > 0) {
    const lengthInches = length * 12;
    const trussesNeeded = Math.ceil(lengthInches / trussSpacing) + 1;

    // Get original truss count (from trussCounts matrix)
    const originalTrusses = lookupMatrix(
      matrices.snow.trussCounts,
      configKey,
      snowCode
    );

    const extraTrusses = Math.max(0, trussesNeeded - originalTrusses);
    if (extraTrusses > 0) {
      const pricePerTruss = matrices.snow.pieTrussPrice || 190;
      totalCost += extraTrusses * pricePerTruss * heightMult;
    }
  }

  // ── Extra Hat Channels ──
  const hatChannelSpacing = lookupMatrix(
    matrices.snow.hatChannelSpacing,
    String(windLoad),
    snowCode
  );
  if (hatChannelSpacing > 0) {
    const halfWidth = (width + 2) / 2; // bar size
    const halfInches = halfWidth * 12;
    const channelsNeeded = Math.ceil(halfInches / hatChannelSpacing) + 1;
    const totalChannels = channelsNeeded * 2; // both sides of roof

    const originalChannels = lookupMatrix(
      matrices.snow.hatChannelCounts,
      String(width),
      "count"
    );

    const extraChannels = Math.max(0, totalChannels - originalChannels);
    if (extraChannels > 0) {
      const pricePerChannel = matrices.snow.channelPricePerFt || 2;
      const channelLength = length;
      totalCost += extraChannels * pricePerChannel * channelLength * heightMult;
    }
  }

  // ── Extra Girts ──
  const girtSpacing = lookupMatrix(
    matrices.snow.girtSpacing,
    String(windLoad),
    String(config.height)
  );
  if (girtSpacing > 0) {
    const heightInches = config.height * 12;
    const girtsNeeded = Math.ceil(heightInches / girtSpacing) + 1;

    const originalGirts = lookupMatrix(
      matrices.snow.girtCounts,
      String(config.height),
      "count"
    );

    const extraGirts = Math.max(0, girtsNeeded - originalGirts);
    if (extraGirts > 0) {
      const pricePerGirt = matrices.snow.tubingPricePerFt || 3;
      totalCost += extraGirts * pricePerGirt * length * heightMult;
    }
  }

  // ── Extra Verticals ──
  const verticalSpacing = lookupMatrix(
    matrices.snow.verticalSpacing,
    String(windLoad),
    String(config.height)
  );
  if (verticalSpacing > 0) {
    const heightInches = config.height * 12;
    const verticalsNeeded = Math.ceil(heightInches / verticalSpacing) + 1;

    const originalVerticals = lookupMatrix(
      matrices.snow.verticalCounts,
      String(width),
      "count"
    );

    const extraVerticals = Math.max(0, verticalsNeeded - originalVerticals);
    if (extraVerticals > 0) {
      const pricePerVertical = matrices.snow.tubingPricePerFt || 3;
      totalCost +=
        extraVerticals * pricePerVertical * config.height * heightMult;
    }
  }

  return Math.round(totalCost);
}

/**
 * Calculate snow/wind engineering costs for widespan buildings.
 *
 * Similar to standard but uses purlin spacing instead of hat channels,
 * and has different truss pricing by width group.
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

  // Wind load (widespan uses 105/120/130)
  const windLoad = config.windRating || 105;

  // Build config key: "E-105-S" or "O-130-G"
  const configKey = `${enclosure}-${windLoad}-${buildingType}`;

  // Snow load code (e.g., "S-30GL", "G-50GL")
  const snowCode = `${buildingType}-${config.snowLoad}`;

  let totalCost = 0;

  // ── Extra Trusses ──
  const trussData = matrices.snow.trussCounts[String(length)];
  if (trussData) {
    const originalTrusses = trussData.count || 0;
    const maxSpacing = matrices.snow.trussSpacing[String(length)]?.spacing || 120;

    const lengthInches = length * 12;
    const trussesNeeded = Math.ceil(lengthInches / maxSpacing) + 1;
    const extraTrusses = Math.max(0, trussesNeeded - originalTrusses);

    if (extraTrusses > 0) {
      const trussPrice =
        lookupValue(matrices.snow.trussPriceByState, String(width)) || 2000;
      totalCost += extraTrusses * trussPrice;

      // Leg height cost for extra trusses
      if (config.height > 10) {
        const legCostPerFt = 90; // per extra foot above base
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

    const originalPurlins = matrices.snow.purlinCounts[String(width)]?.count || 12;
    const extraPurlins = Math.max(0, totalPurlins - originalPurlins);

    if (extraPurlins > 0) {
      const purlinCostPerFt = 6;
      totalCost += extraPurlins * purlinCostPerFt * length;
    }
  }

  // ── Extra Girts ──
  const girtData = matrices.snow.girtSpacing[String(windLoad)];
  if (girtData) {
    const requiredGirtSpacing = girtData.spacing || 60;
    const heightInches = config.height * 12;
    const girtsNeeded = Math.ceil(heightInches / requiredGirtSpacing) + 1;
    const originalGirts = matrices.snow.girtCounts[String(config.height)]?.count || 3;
    const extraGirts = Math.max(0, girtsNeeded - originalGirts);

    if (extraGirts > 0) {
      const perimeter = 2 * (width + length);
      const girtCostPerFt = 6;
      totalCost += extraGirts * girtCostPerFt * perimeter;
    }
  }

  return Math.round(totalCost);
}
