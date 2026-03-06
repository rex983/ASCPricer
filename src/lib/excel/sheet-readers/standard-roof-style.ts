import type { WorkSheet } from "xlsx";
import type { PricingMatrix } from "@/types/pricing";
import { sheetToArray, readMatrix } from "./utils";

/**
 * Parse "Pricing - Roof Style" sheet.
 * Layout: Row 0 = headers ("AFH-12", "AFH-18", ..., "AFV-12", ..., "STD-12", ...)
 *         Column A = lengths (20-100)
 *         Row 1 = zeros (skip)
 *         Data rows 2-18
 *
 * Returns matrix[styleWidth][length] → price surcharge
 * e.g. matrix["AFV-24"]["50"] = 2325
 * STD values are all 0 (base case, no surcharge)
 */
export function readRoofStyle(ws: WorkSheet): PricingMatrix {
  const data = sheetToArray(ws);
  return readMatrix(data, {
    headerRow: 0,
    dataStartRow: 2,
    rowKeyCol: 0,
    dataStartCol: 1,
  });
}
