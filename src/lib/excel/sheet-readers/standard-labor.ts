import type { WorkSheet } from "xlsx";
import type { PricingMatrix } from "@/types/pricing";
import { sheetToArray, readMatrix } from "./utils";

/**
 * Parse "Pricing - Labor-EQ" sheet.
 * Layout: Row 0 = headers ("12S", "18S", ..., "30S", "12T", ..., "30T", "12ET", ..., "30ET")
 *         Column A = lengths (20-100)
 *         Data rows 1-17
 *
 * S = Standard (heights 6-12, often $0)
 * T = Truss/Tube (heights 13-15)
 * ET = Extra Truss (heights 16-20)
 *
 * Returns matrix[widthType][length] → price
 * e.g. matrix["24T"]["50"] = 500
 */
export function readLaborEquipment(ws: WorkSheet): PricingMatrix {
  const data = sheetToArray(ws);
  return readMatrix(data, {
    headerRow: 0,
    dataStartRow: 1,
    rowKeyCol: 0,
    dataStartCol: 1,
  });
}
