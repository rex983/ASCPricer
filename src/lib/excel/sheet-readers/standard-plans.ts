import type { WorkSheet } from "xlsx";
import type { PricingMatrix } from "@/types/pricing";
import { sheetToArray, readMatrix } from "./utils";

/**
 * Parse "Plans for Buildings" sheet.
 * Layout: Row 0 = headers: "Plans", then widths (12, 14, 16, 18, 20, 22, 24, 26, 28, 30)
 *         Column A = lengths (20-100)
 *         Data rows 1-17
 *
 * Returns matrix[width][length] → price
 * e.g. matrix["24"]["50"] = 940
 */
export function readPlans(ws: WorkSheet): PricingMatrix {
  const data = sheetToArray(ws);
  return readMatrix(data, {
    headerRow: 0,
    dataStartRow: 1,
    rowKeyCol: 0,
    dataStartCol: 1,
    dataEndCol: 12, // stop before Calcs section
  });
}
