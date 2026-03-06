import type { WorkSheet } from "xlsx";
import type { PricingMatrix } from "@/types/pricing";
import { sheetToArray, readMatrix } from "./utils";

/**
 * Parse widespan "Equipment" sheet.
 * Layout: Row 0 = headers: lengths (20, 25, 30, ..., 200)
 *         Column A = widths (32, 34, 36, ..., 60)
 *         Row 1 = zeros (skip)
 *         Data rows 2-16
 *
 * Three pricing tiers by length:
 *   Widths 32-46: 2300 (20-60), 3300 (65-100), 4600 (105+)
 *   Widths 48-60: 3850 (20-60), 4350 (65-100), 5150 (105+)
 *
 * Returns matrix[width][length] → price
 */
export function readWidespanEquipment(ws: WorkSheet): PricingMatrix {
  const data = sheetToArray(ws);
  return readMatrix(data, {
    headerRow: 0,
    dataStartRow: 2,
    rowKeyCol: 0,
    dataStartCol: 1,
    transpose: true, // matrix[width][length]
  });
}
