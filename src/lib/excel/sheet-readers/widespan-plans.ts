import type { WorkSheet } from "xlsx";
import type { PricingMatrix } from "@/types/pricing";
import { sheetToArray, readMatrix } from "./utils";

/**
 * Parse widespan "Plans & Calcs" sheet.
 * Layout: Row 0 = headers: widths (32, 34, 36, ..., 60)
 *         Column A = lengths (20-130)
 *         Data rows 2-24
 *
 * Returns matrix[width][length] → price
 */
export function readWidespanPlans(ws: WorkSheet): PricingMatrix {
  const data = sheetToArray(ws);
  return readMatrix(data, {
    headerRow: 0,
    dataStartRow: 2,
    rowKeyCol: 0,
    dataStartCol: 1,
    dataEndCol: 17, // stop before Calcs section
  });
}
