import type { WorkSheet } from "xlsx";
import type { PricingLookup } from "@/types/pricing";
import { sheetToArray, num, cleanHeader } from "./utils";

/**
 * Parse widespan "Insulation - Wainscot" sheet.
 * Insulation is formula-based (rate × sqft), not a matrix lookup.
 * We extract the rates and the wainscot lookup tables.
 *
 * Layout:
 *   Row 0, col 16: "Sides" header, col 19: "Ends" header
 *   Rows 1+:
 *     col 15 = length key, col 16 = sides price
 *     col 18 = width key, col 19 = ends price
 */
export function readWidespanInsulationWainscot(ws: WorkSheet): {
  fiberglassRate: number;
  thermalRate: number;
  wainscotSides: PricingLookup;
  wainscotEnds: PricingLookup;
} {
  const data = sheetToArray(ws);

  let fiberglassRate = 2.25;
  let thermalRate = 1.65;

  // Scan for insulation rates in the left side of the sheet
  for (let r = 0; r < Math.min(10, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 0; c < Math.min(14, row.length); c++) {
      const label = cleanHeader(row[c]);
      if (label.toLowerCase().includes("fiberglass") || label.toLowerCase().includes("fiber")) {
        for (let cc = c - 2; cc <= c + 2; cc++) {
          if (cc >= 0 && cc < row.length) {
            const v = num(row[cc]);
            if (v > 1 && v < 5) fiberglassRate = v;
          }
        }
      }
      if (label.toLowerCase().includes("prodex") || label.toLowerCase().includes("thermal")) {
        for (let cc = c - 2; cc <= c + 2; cc++) {
          if (cc >= 0 && cc < row.length) {
            const v = num(row[cc]);
            if (v > 1 && v < 5) thermalRate = v;
          }
        }
      }
    }
  }

  // Find the wainscot table columns by scanning row 0 for "Sides" and "Ends" headers
  let sidesKeyCol = 15;
  let sidesValCol = 16;
  let endsKeyCol = 18;
  let endsValCol = 19;

  const headerRow = data[0] || [];
  for (let c = 10; c < headerRow.length; c++) {
    const h = cleanHeader(headerRow[c]);
    if (h === "Sides") {
      sidesKeyCol = c - 1;
      sidesValCol = c;
    }
    if (h === "Ends") {
      endsKeyCol = c - 1;
      endsValCol = c;
    }
  }

  // Read wainscot sides (length → price)
  const wainscotSides: PricingLookup = {};
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row) continue;
    const key = num(row[sidesKeyCol]);
    const val = num(row[sidesValCol]);
    if (key <= 0 || val <= 0) break;
    wainscotSides[String(key)] = val;
  }

  // Read wainscot ends (width → price)
  const wainscotEnds: PricingLookup = {};
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row) continue;
    const key = num(row[endsKeyCol]);
    const val = num(row[endsValCol]);
    if (key <= 0 || val <= 0) break;
    wainscotEnds[String(key)] = val;
  }

  return { fiberglassRate, thermalRate, wainscotSides, wainscotEnds };
}
