import type { WorkSheet } from "xlsx";
import type { PricingLookup } from "@/types/pricing";
import { sheetToArray, num, cleanHeader } from "./utils";

export interface StandardChangersResult {
  widthBuckets: PricingLookup;
  lengthBuckets: PricingLookup;
  heightToSidesKey: PricingLookup;
  heightToLegsKey: PricingLookup;
  gaugeLookup: PricingLookup;
  sheetMetalMultiplier: PricingLookup;
  buildingTypeByHeight: PricingLookup;
}

/**
 * Parse "Pricing - Changers" sheet.
 * This sheet has multiple sections stacked vertically:
 *
 * R0-R8: Width changers - maps input widths to column buckets
 * R9-R17: Length changers - maps input lengths to row buckets
 * R18-R24: Gauge changers - normalizes gauge input (12G, 14G, etc.)
 * R26-R33: Roof Style changers (not stored, resolved at runtime)
 * R35-R42: Sides price mapping - height → sides lookup key
 * R44-R49: Leg price mapping - height → legs lookup key
 *
 * Each section has a mapping row where col indices → resolved values.
 * We extract the mapping as a lookup table.
 */
export function readChangers(ws: WorkSheet): StandardChangersResult {
  const data = sheetToArray(ws);

  // Width buckets: Row 1 maps indices to width values
  const widthBuckets = extractBucketMapping(data, 1);

  // Length buckets: Row 10 maps indices to length values
  const lengthBuckets = extractBucketMapping(data, 10);

  // Gauge lookup: Row 18 = input keys, Row 19 = normalized values
  const gaugeLookup = extractKeyValueMapping(data, 18, 19);

  // Sides height mapping: Row 35 = input heights, Row 36 = sides lookup keys
  const heightToSidesKey = extractKeyValueMapping(data, 35, 36);

  // Legs height mapping: Row 44 = input heights, Row 45 = legs lookup keys
  const heightToLegsKey = extractKeyValueMapping(data, 44, 45);

  // Building type by height: derive from Labor-EQ lookup area (rows 20-24)
  // Heights 6-12 → "S", 13-15 → "T", 16-20 → "ET"
  const buildingTypeByHeight: PricingLookup = {};
  for (let h = 6; h <= 12; h++) buildingTypeByHeight[String(h)] = 0; // S
  for (let h = 13; h <= 15; h++) buildingTypeByHeight[String(h)] = 1; // T
  for (let h = 16; h <= 20; h++) buildingTypeByHeight[String(h)] = 2; // ET

  // Sheet metal multiplier (standard only has one gauge option, but we include for consistency)
  const sheetMetalMultiplier: PricingLookup = {
    "29G Agg": 1.0,
    "26G Agg": 1.1,
    "26G PBR": 1.2,
  };

  return {
    widthBuckets,
    lengthBuckets,
    heightToSidesKey,
    heightToLegsKey,
    gaugeLookup,
    sheetMetalMultiplier,
    buildingTypeByHeight,
  };
}

/**
 * Extract a bucket mapping from a row where sequential columns map to bucket values.
 * The row contains values like: [_, 12, 12, 12, 18, 18, 20, 22, 24, ...]
 * We invert this to: { "12": 12, "18": 18, ... } (unique values only needed)
 * But more usefully, we want: input value → nearest bucket.
 */
function extractBucketMapping(
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

/**
 * Extract a key-value mapping from two rows.
 * Row keyRow has input values, row valueRow has resolved values.
 */
function extractKeyValueMapping(
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
