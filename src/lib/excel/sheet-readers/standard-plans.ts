import type { WorkSheet } from "xlsx";
import type { PricingMatrix, PricingLookup } from "@/types/pricing";
import { sheetToArray, readMatrix, readLookup } from "./utils";

/**
 * Parse "Plans for Buildings" sheet.
 * Layout: Row 0 = headers: "Plans", then widths (12-30) | gap | "Calcs", then widths (12-30)
 *         Column A = lengths (20-100)
 *         Data rows 1-17
 *
 * Plans section: columns 1-11 → matrix[width][length] → base plans price
 * Calcs section: columns 15-25 → matrix[width][length] → calculations price
 *
 * Also has surcharge tables below the main matrices:
 * - Leg height surcharge: rows 27-41, col B=height, col C=surcharge
 * - Door opening cost: rows 34-46, col K=doorCount, col L=cost
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

/** Parse the Calculations matrix (right side of the "Plans for Buildings" sheet) */
export function readCalculations(ws: WorkSheet): PricingMatrix {
  const data = sheetToArray(ws);
  return readMatrix(data, {
    headerRow: 0,
    dataStartRow: 1,
    rowKeyCol: 0, // lengths in col A (same as plans)
    dataStartCol: 16, // Skip "Calcs" label at col 15; width headers start at col 16
    dataEndCol: 27, // widths 12-30 (11 columns: 16-26)
  });
}

/** Parse the leg height surcharge lookup (rows 27-41, col B=height, col C=surcharge) */
export function readPlansLegSurcharge(ws: WorkSheet): PricingLookup {
  const data = sheetToArray(ws);
  return readLookup(data, {
    startRow: 27,
    endRow: 42,
    keyCol: 1, // col B = height
    valueCol: 2, // col C = surcharge
  });
}

/** Parse the door opening cost lookup (rows 34-46, col K=doorCount, col L=cost) */
export function readPlansDoorOpeningCost(ws: WorkSheet): PricingLookup {
  const data = sheetToArray(ws);
  return readLookup(data, {
    startRow: 34,
    endRow: 47,
    keyCol: 10, // col K = door count
    valueCol: 11, // col L = cost
  });
}
