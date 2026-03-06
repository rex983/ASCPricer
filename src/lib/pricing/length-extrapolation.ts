import { LENGTH_DECOMPOSITION } from "./constants";
import type { PricingMatrix } from "@/types/pricing";
import { lookupMatrix } from "./lookups";

/**
 * For standard buildings, lengths 55-100 are computed by summing
 * two base-length prices. This function handles that decomposition.
 *
 * For lengths 20-50, it returns the direct lookup value.
 * For lengths 55-100, it sums the two component lengths.
 */
export function lookupWithLengthExtrapolation(
  matrix: PricingMatrix,
  rowKey: string,
  length: number
): number {
  const decomposition = LENGTH_DECOMPOSITION[length];

  if (!decomposition) {
    // Direct lookup for base lengths (20-50)
    return lookupMatrix(matrix, rowKey, String(length));
  }

  const [len1, len2] = decomposition;
  const price1 = lookupMatrix(matrix, rowKey, String(len1));
  const price2 = lookupMatrix(matrix, rowKey, String(len2));

  return price1 + price2;
}
