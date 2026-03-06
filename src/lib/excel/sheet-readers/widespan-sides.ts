import type { WorkSheet } from "xlsx";
import type { PricingMatrix } from "@/types/pricing";
import { sheetToArray, readMatrix } from "./utils";

/**
 * Parse widespan "Sides" sheet.
 * Layout: Row 0 = headers: lengths (20, 25, 30, ..., 200)
 *         Column A = number of enclosed side panels (3-16)
 *         Data rows 2-15
 *
 * Returns matrix[panelCount][length] → price
 */
export function readWidespanSides(ws: WorkSheet): PricingMatrix {
  const data = sheetToArray(ws);
  return readMatrix(data, {
    headerRow: 0,
    dataStartRow: 2,
    rowKeyCol: 0,
    dataStartCol: 1,
    transpose: true, // matrix[panelCount][length]
  });
}
