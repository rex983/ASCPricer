import type { WorkSheet } from "xlsx";
import type { PricingMatrix } from "@/types/pricing";
import { sheetToArray, num, cleanHeader } from "./utils";

/**
 * Parse "Pricing - Ends" sheet.
 * Layout:
 *   Row 0 = headers: "12-HZ-FE", "18-HZ-FE", ..., "12-V-FE", ..., "12-HZ-G", ..., "12-V-G", ..., "12-HZ-EG", ...
 *   Column A = heights (6-20)
 *   Row 1 = zeros (skip)
 *   Data rows 2-16
 *
 * Also has V-end surcharge tables at rows 18-23
 *
 * Returns:
 *   ends: matrix[widthOrientationType][height] → price
 *   vEndsSurcharge: matrix for V-end upcharges
 */
export function readEnds(ws: WorkSheet): {
  ends: PricingMatrix;
  vEndsSurcharge: PricingMatrix;
} {
  const data = sheetToArray(ws);
  const headers = data[0] || [];

  // Main ends matrix
  const ends: PricingMatrix = {};

  for (let r = 2; r < Math.min(17, data.length); r++) {
    const row = data[r];
    if (!row) break;
    const height = String(num(row[0]));
    if (height === "0" || num(height) < 6) continue;

    for (let c = 1; c < headers.length; c++) {
      const colKey = cleanHeader(headers[c]);
      if (!colKey || colKey === "0") continue;
      if (!ends[colKey]) ends[colKey] = {};
      ends[colKey][height] = num(row[c]);
    }
  }

  // V-end surcharge tables (rows 18+)
  const vEndsSurcharge: PricingMatrix = {};

  for (let r = 17; r < Math.min(30, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    const label = cleanHeader(row[0]);
    if (!label) continue;

    if (
      label.includes("V Ends") ||
      label.includes("VEADD") ||
      label.includes("V ends")
    ) {
      vEndsSurcharge[label] = {};
      for (let c = 1; c < headers.length; c++) {
        const colKey = cleanHeader(headers[c]);
        if (!colKey || colKey === "0") continue;
        // Extract width from the header "XX-HZ-FE" → "XX"
        const widthMatch = colKey.match(/^(\d+)/);
        if (widthMatch) {
          vEndsSurcharge[label][widthMatch[1]] = num(row[c]);
        }
      }
    }
  }

  return { ends, vEndsSurcharge };
}
