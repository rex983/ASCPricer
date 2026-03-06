import type { WorkSheet } from "xlsx";
import type { PricingMatrix } from "@/types/pricing";
import { sheetToArray, readMatrix } from "./utils";

/**
 * Parse widespan "Base Price" sheet.
 * Layout: Row 0 = headers: width-gauge pairs ("32-14G", "34-14G", ..., "60-14G", "32-12G", ..., "60-12G")
 *         Column A = lengths (20-200 in 5ft increments)
 *         Row 1 = zeros (skip)
 *         Data rows 2-38
 *
 * Returns matrix[widthGauge][length] → price
 * e.g. matrix["60-12G"]["50"] = 15740
 */
export function readWidespanBasePrice(ws: WorkSheet): PricingMatrix {
  const data = sheetToArray(ws);
  return readMatrix(data, {
    headerRow: 0,
    dataStartRow: 2,
    rowKeyCol: 0,
    dataStartCol: 1,
  });
}
