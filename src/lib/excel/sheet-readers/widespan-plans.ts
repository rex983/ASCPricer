import type { WorkSheet } from "xlsx";
import type { PricingMatrix } from "@/types/pricing";
import { sheetToArray, readMatrix } from "./utils";

/**
 * Parse widespan "Plans & Calcs" sheet.
 * Layout: Row 0 = headers: widths (32, 34, 36, ..., 60)
 *         Column A = lengths (20-130)
 *         Plans section: cols 1-16 (A:Q)
 *         Calcs section: cols 18+ (S:AI)
 *
 * Returns { plans, calculations } — both matrix[width][length] → price
 */
export function readWidespanPlans(ws: WorkSheet): {
  plans: PricingMatrix;
  calculations: PricingMatrix;
} {
  const data = sheetToArray(ws);

  // Plans matrix: cols 1-16
  const plans = readMatrix(data, {
    headerRow: 0,
    dataStartRow: 2,
    rowKeyCol: 0,
    dataStartCol: 1,
    dataEndCol: 17, // stop before Calcs section
  });

  // Calculations matrix: cols 18+ (same row layout, different column range)
  const calculations = readMatrix(data, {
    headerRow: 0,
    dataStartRow: 2,
    rowKeyCol: 0,
    dataStartCol: 18,
    dataEndCol: 35, // S through AI (columns 18-34)
  });

  return { plans, calculations };
}
