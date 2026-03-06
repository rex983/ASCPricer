import type { WorkSheet } from "xlsx";
import type { PricingMatrix } from "@/types/pricing";
import { sheetToArray, readMatrix } from "./utils";

/**
 * Parse widespan "Leg Height" sheet.
 * Layout: Row 0 = headers: lengths (20, 25, 30, ..., 200)
 *         Column A = leg heights (8-20)
 *         Row 1 = zeros (skip)
 *         Data rows 2-14
 *         Heights 8-10 are all zeros (included in base price)
 *         Upcharges start at height 11
 *
 * Returns matrix[height][length] → price
 */
export function readWidespanLegs(ws: WorkSheet): PricingMatrix {
  const data = sheetToArray(ws);
  return readMatrix(data, {
    headerRow: 0,
    dataStartRow: 2,
    rowKeyCol: 0,
    dataStartCol: 1,
    transpose: true, // matrix[height][length]
  });
}
