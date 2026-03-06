import type { WorkSheet } from "xlsx";
import type { PricingLookup } from "@/types/pricing";
import { sheetToArray, num, cleanHeader } from "./utils";

/**
 * Parse "Pricing - Accessories" sheet.
 * Layout: Two price lists side by side:
 *   Windows (col A-B, rows 0-8): name → price
 *   Doors (col H-I, rows 0-8): name → price
 *
 * Returns { walkInDoors, windows }
 */
export function readAccessories(ws: WorkSheet): {
  walkInDoors: PricingLookup;
  windows: PricingLookup;
} {
  const data = sheetToArray(ws);

  const windows: PricingLookup = {};
  const walkInDoors: PricingLookup = {};

  // Read windows from cols A-B (0-1)
  for (let r = 0; r < Math.min(20, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    const name = cleanHeader(row[0]);
    const price = num(row[1]);
    if (name && price > 0 && !name.toLowerCase().includes("price") && !name.toLowerCase().includes("total")) {
      windows[name] = price;
    }
  }

  // Read doors from cols further right (typically col 7-8 or wherever "Door" items start)
  // Scan to find the door column
  for (let c = 2; c < 15; c++) {
    for (let r = 0; r < Math.min(20, data.length); r++) {
      const row = data[r];
      if (!row) continue;
      const name = cleanHeader(row[c]);
      if (name && name.toLowerCase().includes("door")) {
        // Found door column - read all entries from this column pair
        for (let dr = 0; dr < Math.min(20, data.length); dr++) {
          const drow = data[dr];
          if (!drow) continue;
          const dname = cleanHeader(drow[c]);
          const dprice = num(drow[c + 1]);
          if (dname && dprice > 0 && !dname.toLowerCase().includes("price") && !dname.toLowerCase().includes("total")) {
            walkInDoors[dname] = dprice;
          }
        }
        return { walkInDoors, windows };
      }
    }
  }

  return { walkInDoors, windows };
}

/**
 * Parse standard roll-up door pricing.
 * Standard spreadsheet doesn't have a separate roll-up sheet,
 * so roll-ups may be in the Accessories sheet or Quote Sheet.
 * Returns empty lookup if not found - will be populated from widespan pattern.
 */
export function readStandardRollUpDoors(ws: WorkSheet): PricingLookup {
  const data = sheetToArray(ws);
  const rollUps: PricingLookup = {};

  // Scan for roll-up door entries
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 0; c < row.length - 1; c++) {
      const name = cleanHeader(row[c]);
      if (name && (name.toLowerCase().includes("roll") || name.match(/^\d+x\d+$/))) {
        const price = num(row[c + 1]);
        if (price > 0) {
          rollUps[name] = price;
        }
      }
    }
  }

  return rollUps;
}
