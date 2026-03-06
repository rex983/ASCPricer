import type { PricingMatrix, PricingLookup } from "@/types/pricing";

/**
 * 2D matrix lookup — equivalent to INDEX/MATCH in the spreadsheet.
 * Returns 0 if the key combination is not found.
 */
export function lookupMatrix(
  matrix: PricingMatrix,
  rowKey: string,
  colKey: string
): number {
  return matrix[rowKey]?.[colKey] ?? 0;
}

/**
 * 1D lookup — equivalent to VLOOKUP in the spreadsheet.
 */
export function lookupValue(
  lookup: PricingLookup,
  key: string
): number {
  return lookup[key] ?? 0;
}

/**
 * Find the nearest bucket for a given value from an array of valid buckets.
 * Rounds down to the nearest valid bucket.
 */
export function nearestBucket(value: number, buckets: readonly number[]): number {
  let nearest = buckets[0];
  for (const bucket of buckets) {
    if (bucket <= value) {
      nearest = bucket;
    } else {
      break;
    }
  }
  return nearest;
}

/**
 * Round a value to the nearest increment (e.g., nearest 5ft for length).
 */
export function roundToIncrement(value: number, increment: number): number {
  return Math.round(value / increment) * increment;
}
