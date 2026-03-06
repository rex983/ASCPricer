import type { WorkSheet } from "xlsx";
import type { PricingMatrix } from "@/types/pricing";
import { sheetToArray, num, cleanHeader } from "./utils";

/**
 * Parse "Pricing - Legs" sheet.
 * Layout: Two side-by-side matrices:
 *   Block 1 (cols 1-18): widths 12-24, headers are lengths (20-100)
 *   Block 2 (cols 20-37): widths 26-30, headers are lengths (20-100)
 *   Column A = heights (6-20)
 *   Row 0 = length headers
 *   Row 1 = zeros (skip)
 *   Data rows 2-16
 *
 * Returns { small, large } where each is matrix[height][length] → price
 */
export function readLegs(ws: WorkSheet): {
  small: PricingMatrix;
  large: PricingMatrix;
} {
  const data = sheetToArray(ws);
  const headers = data[0] || [];

  const small: PricingMatrix = {};
  const large: PricingMatrix = {};

  // Find the separator between the two blocks (empty column)
  let block2Start = -1;
  for (let c = 2; c < headers.length; c++) {
    const h = cleanHeader(headers[c]);
    if (h === "" || h === "0") {
      // Check if next column restarts with a length header
      const next = cleanHeader(headers[c + 1]);
      if (next && num(next) > 0) {
        block2Start = c + 1;
        break;
      }
    }
  }

  for (let r = 2; r < data.length; r++) {
    const row = data[r];
    if (!row) break;
    const height = cleanHeader(row[0]);
    if (!height || height === "0") continue;
    const heightNum = num(height);
    if (heightNum < 6 || heightNum > 20) continue;

    const heightKey = String(heightNum);
    small[heightKey] = {};
    if (block2Start > 0) large[heightKey] = {};

    // Block 1: widths 12-24
    for (let c = 1; c < (block2Start > 0 ? block2Start - 1 : headers.length); c++) {
      const colKey = cleanHeader(headers[c]);
      if (!colKey || colKey === "0") continue;
      small[heightKey][colKey] = num(row[c]);
    }

    // Block 2: widths 26-30
    if (block2Start > 0) {
      for (let c = block2Start; c < headers.length; c++) {
        const colKey = cleanHeader(headers[c]);
        if (!colKey || colKey === "0") continue;
        large[heightKey][colKey] = num(row[c]);
      }
    }
  }

  return { small, large };
}
