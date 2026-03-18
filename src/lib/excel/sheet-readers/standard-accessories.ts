import type { WorkSheet } from "xlsx";
import type { PricingLookup } from "@/types/pricing";
import { sheetToArray, num, cleanHeader } from "./utils";

/**
 * Parse "Pricing - Accessories" sheet.
 * Layout: Two price lists side by side:
 *   Windows (col A-B, rows 0-8): name → price
 *   Doors (col H-I or similar, rows 0-10): name → price
 *
 * Scalable: scans for name/price column pairs using content patterns
 * rather than hardcoded column positions.
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

  // Read windows from cols A-B (0-1) — always in this position
  for (let r = 0; r < Math.min(20, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    const name = cleanHeader(row[0]);
    const price = num(row[1]);
    if (name && price > 0 && !name.toLowerCase().includes("price") && !name.toLowerCase().includes("total")) {
      windows[name] = price;
    }
  }

  // Find the door column by scanning for name/price pairs that look like doors.
  // Door names contain dimension patterns (e.g., 36"x80") or keywords like
  // "Swing", "Panel", "Frame Out", "Lock", "Lite", "Buck", "Diamond".
  const doorPatterns = [/\d+"?\s*x\s*\d+"?/, /swing/i, /panel/i, /frame\s*out/i, /lock/i, /lite/i, /buck/i, /diamond/i];

  const isDoorName = (name: string) =>
    doorPatterns.some((p) => p.test(name));

  for (let c = 2; c < 15; c++) {
    // Count how many rows in this column look like door entries
    let doorHits = 0;
    for (let r = 0; r < Math.min(15, data.length); r++) {
      const row = data[r];
      if (!row) continue;
      const name = cleanHeader(row[c]);
      const price = num(row[c + 1]);
      if (name && price > 0 && isDoorName(name)) doorHits++;
    }

    if (doorHits >= 2) {
      // This column has door data — read all name/price pairs
      for (let r = 0; r < Math.min(20, data.length); r++) {
        const row = data[r];
        if (!row) continue;
        const name = cleanHeader(row[c]);
        const price = num(row[c + 1]);
        if (
          name &&
          price > 0 &&
          !name.toLowerCase().includes("price") &&
          !name.toLowerCase().includes("total") &&
          !name.toLowerCase().includes("qty")
        ) {
          walkInDoors[name] = price;
        }
      }
      break;
    }
  }

  return { walkInDoors, windows };
}

/**
 * Parse standard roll-up door pricing from "Pricing - Accessories" sheet.
 *
 * Layout: Two separate price columns for roll-ups:
 *   Col P-Q (15-16): size → ENDS price (e.g., 8x8 = $795)
 *   Col S-T (18-19): size → SIDES price (e.g., 8x8 = $1,055)
 *
 * Returns { rollUpEnds, rollUpSides }
 */
export function readStandardRollUpDoors(ws: WorkSheet): {
  rollUpEnds: PricingLookup;
  rollUpSides: PricingLookup;
} {
  const data = sheetToArray(ws);
  const rollUpEnds: PricingLookup = {};
  const rollUpSides: PricingLookup = {};

  // Ends prices: col P (15) = size, col Q (16) = price
  for (let r = 0; r < Math.min(25, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    const size = cleanHeader(row[15]);
    const price = num(row[16]);
    if (size && size.match(/^\d+x\d+$/) && price > 0) {
      rollUpEnds[size] = price;
    }
  }

  // Sides prices: col S (18) = size, col T (19) = price
  for (let r = 0; r < Math.min(25, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    const size = cleanHeader(row[18]);
    const price = num(row[19]);
    if (size && size.match(/^\d+x\d+$/) && price > 0) {
      rollUpSides[size] = price;
    }
  }

  return { rollUpEnds, rollUpSides };
}
