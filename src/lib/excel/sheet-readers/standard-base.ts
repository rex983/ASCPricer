import type { WorkSheet } from "xlsx";
import type { PricingMatrix } from "@/types/pricing";
import { sheetToArray, readMatrix } from "./utils";

/**
 * Parse "Pricing - Base" sheet.
 * Layout: Row 0 = headers (width-gauge: "12-14G", "18-14G", ..., "30-12G")
 *         Column A = lengths (20, 25, 30, ... 100)
 *         Row 1 = zeros (skip)
 *         Data rows 2-18
 *
 * Returns matrix[widthGauge][length] → price
 * e.g. matrix["24-14G"]["50"] = 4790
 */
export function readBasePrice(ws: WorkSheet): PricingMatrix {
  const data = sheetToArray(ws);
  return readMatrix(data, {
    headerRow: 0,
    dataStartRow: 2, // skip zero row
    rowKeyCol: 0,
    dataStartCol: 1,
  });
}
