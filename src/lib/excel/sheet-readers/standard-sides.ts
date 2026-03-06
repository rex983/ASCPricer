import type { WorkSheet } from "xlsx";
import type { PricingMatrix } from "@/types/pricing";
import { sheetToArray, num, cleanHeader } from "./utils";

/**
 * Parse "Pricing - Sides" sheet.
 * Layout:
 *   Row 0 = headers: "0-HZ", "20-HZ", "25-HZ", ..., "100-HZ", "0-V", "20-V", ..., "100-V"
 *   Rows 1-19 = side heights: "0' Sides Down", "3' Sides Down", ..., "20' Sides Down"
 *   Column A = side height labels
 *
 * Also has V-side surcharge tables at rows 22-27:
 *   "6'-10' V Sides" and "11'-15' V Sides" surcharges by length
 *   "VsADD 16-20" additional surcharge
 *
 * Returns:
 *   sides: matrix[lengthOrientation][sideHeight] → price
 *   vSidesSurcharge: matrix for V-side upcharges
 */
export function readSides(ws: WorkSheet): {
  sides: PricingMatrix;
  vSidesSurcharge: PricingMatrix;
} {
  const data = sheetToArray(ws);
  const headers = data[0] || [];

  // Main sides matrix
  const sides: PricingMatrix = {};

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row) break;
    const rawLabel = cleanHeader(row[0]);
    if (!rawLabel || rawLabel === "0") continue;

    // Extract the numeric height from "X' Sides Down"
    const match = rawLabel.match(/^(\d+)'/);
    if (!match && !rawLabel.includes("0'")) continue;
    const sideHeight = match ? match[1] : "0";

    for (let c = 1; c < headers.length; c++) {
      const colKey = cleanHeader(headers[c]);
      if (!colKey || colKey === "0") continue;
      if (!sides[colKey]) sides[colKey] = {};
      sides[colKey][sideHeight] = num(row[c]);
    }

    // Stop at the blank row before surcharge tables
    if (r >= 20) break;
  }

  // V-side surcharge tables (rows after main matrix)
  const vSidesSurcharge: PricingMatrix = {};
  let surchargeSection = false;

  for (let r = 20; r < Math.min(30, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    const label = cleanHeader(row[0]);
    if (!label) continue;

    // Look for V Sides surcharge rows
    if (
      label.includes("V Sides") ||
      label.includes("VsADD") ||
      label.includes("V sides")
    ) {
      surchargeSection = true;
      vSidesSurcharge[label] = {};
      for (let c = 1; c < headers.length; c++) {
        const colKey = cleanHeader(headers[c]);
        if (!colKey || colKey === "0") continue;
        // Extract just the length from "XX-HZ" or "XX-V"
        const lengthMatch = colKey.match(/^(\d+)/);
        if (lengthMatch) {
          vSidesSurcharge[label][lengthMatch[1]] = num(row[c]);
        }
      }
    }
  }

  return { sides, vSidesSurcharge };
}
