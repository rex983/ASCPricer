import type { WorkSheet } from "xlsx";
import type { PricingMatrix } from "@/types/pricing";
import { sheetToArray, readMatrix } from "./utils";

/**
 * Parse widespan "Ends" sheet.
 * Layout: Row 0 = headers: "{width}-{endType}" (e.g., "32-FE", "34-FE", ..., "60-FE", "32-GE", ..., "60-GE")
 *         Column A = leg heights (10-20)
 *         Data rows 1-11
 *         FE = Full Enclosed (varies by height)
 *         GE = Gabled End (flat per width)
 *         0 = Open (all zeros)
 *
 * Returns matrix[widthEndType][height] → price
 */
export function readWidespanEnds(ws: WorkSheet): PricingMatrix {
  const data = sheetToArray(ws);
  return readMatrix(data, {
    headerRow: 0,
    dataStartRow: 1,
    rowKeyCol: 0,
    dataStartCol: 1,
  });
}
