import type { WorkSheet } from "xlsx";
import type { PricingLookup, PricingMatrix } from "@/types/pricing";
import { sheetToArray, num, cleanHeader } from "./utils";

export interface WidespanChangersResult {
  widthBuckets: PricingLookup;
  lengthBuckets: PricingLookup;
  gaugeLookup: PricingLookup;
  sheetMetalMultiplier: PricingLookup;
  buildingTypeByWidthHeight: PricingMatrix;
  windLoadMapping: PricingLookup;
  snowLoadMapping: PricingLookup;
  heightClassification: PricingLookup;
}

/**
 * Parse widespan "Changers" sheet.
 * This wide sheet (202 cols) has multiple stacked sections:
 *
 * R0-R7: Length changer - maps input length to nearest 5ft bucket
 * R8-R14: Width changer - maps input width to nearest 2ft bucket
 * R16-R21: Gauge/Framing - normalizes gauge (12G, 14G)
 * R24-R29: Sheet Metal - maps panel type to multiplier (1.0, 1.1, 1.2)
 * R31-R35: Width → Building Type (S for 32-40, G for 42-60)
 * R37-R42: Snow Load mapping
 * R44-R49: Wind Load mapping (first table)
 * R51-R56: Height → Building Type (S for 8-12, G for 13-20)
 * R58-R63: Wind Load mapping (second table - don't use)
 * R65-R70: Height normalization (min 10)
 */
export function readWidespanChangers(ws: WorkSheet): WidespanChangersResult {
  const data = sheetToArray(ws);

  // Length buckets (row 1 has the mapping values)
  const lengthBuckets = extractUniqueBuckets(data, 1);

  // Width buckets (row 9)
  const widthBuckets = extractUniqueBuckets(data, 9);

  // Gauge lookup (row 16 = input keys, row 17 = normalized)
  const gaugeLookup = extractPairMapping(data, 16, 17);

  // Sheet metal multiplier (row 24 = panel type names, row 25 = multipliers)
  const sheetMetalMultiplier = extractPairMappingStr(data, 24, 25);

  // Building type by width (row 31 = width indices, row 32 = S/G)
  // Building type by height (row 51 = height indices, row 52 = S/G)
  const buildingTypeByWidthHeight: PricingMatrix = {};

  // Width → type: S for 32-40, G for 42-60
  const widthTypeRow = data[32] || data[33];
  const widthKeyRow = data[31] || data[32];
  if (widthKeyRow && widthTypeRow) {
    for (let c = 1; c < Math.min(widthKeyRow.length, widthTypeRow.length); c++) {
      const width = String(num(widthKeyRow[c]));
      const type = cleanHeader(widthTypeRow[c]);
      if (width !== "0" && type) {
        if (!buildingTypeByWidthHeight[width]) buildingTypeByWidthHeight[width] = {};
        buildingTypeByWidthHeight[width]["type"] = type === "S" ? 0 : 1;
      }
    }
  }

  // Height → type: S for 8-12, G for 13-20
  const heightTypeRow = data[53] || data[54];
  const heightKeyRow = data[52] || data[53];
  if (heightKeyRow && heightTypeRow) {
    for (let c = 1; c < Math.min(heightKeyRow.length, heightTypeRow.length); c++) {
      const height = String(num(heightKeyRow[c]));
      const type = cleanHeader(heightTypeRow[c]);
      if (height !== "0" && type) {
        if (!buildingTypeByWidthHeight[height])
          buildingTypeByWidthHeight[height] = {};
        buildingTypeByWidthHeight[height]["heightType"] = type === "S" ? 0 : 1;
      }
    }
  }

  // Snow load mapping (row 37 = descriptions, row 38 = codes)
  const snowLoadMapping = extractPairMappingStr(data, 37, 38);

  // Wind load mapping (row 44 = input, row 45 = resolved)
  const windLoadMapping = extractPairMapping(data, 44, 45);

  // Height classification (row 65 = input, row 66 = normalized min 10)
  const heightClassification = extractPairMapping(data, 65, 66);

  return {
    widthBuckets,
    lengthBuckets,
    gaugeLookup,
    sheetMetalMultiplier,
    buildingTypeByWidthHeight,
    windLoadMapping,
    snowLoadMapping,
    heightClassification,
  };
}

function extractUniqueBuckets(
  data: (string | number)[][],
  row: number
): PricingLookup {
  const rowData = data[row];
  if (!rowData) return {};
  const buckets: PricingLookup = {};
  const seen = new Set<number>();
  for (let c = 1; c < rowData.length; c++) {
    const val = num(rowData[c]);
    if (val > 0 && !seen.has(val)) {
      seen.add(val);
      buckets[String(val)] = val;
    }
  }
  return buckets;
}

function extractPairMapping(
  data: (string | number)[][],
  keyRow: number,
  valueRow: number
): PricingLookup {
  const keys = data[keyRow];
  const values = data[valueRow];
  if (!keys || !values) return {};
  const lookup: PricingLookup = {};
  for (let c = 1; c < Math.min(keys.length, values.length); c++) {
    const key = String(num(keys[c]));
    const val = num(values[c]);
    if (key !== "0" && val > 0) {
      lookup[key] = val;
    }
  }
  return lookup;
}

function extractPairMappingStr(
  data: (string | number)[][],
  keyRow: number,
  valueRow: number
): PricingLookup {
  const keys = data[keyRow];
  const values = data[valueRow];
  if (!keys || !values) return {};
  const lookup: PricingLookup = {};
  for (let c = 1; c < Math.min(keys.length, values.length); c++) {
    const key = cleanHeader(keys[c]);
    const val = num(values[c]);
    if (key && key !== "0") {
      lookup[key] = val;
    }
  }
  return lookup;
}
