import type { WorkSheet } from "xlsx";
import type { PricingLookup } from "@/types/pricing";
import { sheetToArray, num, cleanHeader } from "./utils";

/**
 * Parse widespan "Doors - Windows" sheet.
 * Layout: Doors section (rows 0-9), then Windows section (rows 16-24)
 *   Column A = item name, Column B = price
 *
 * Returns { walkInDoors, windows }
 */
export function readWidespanDoorsWindows(ws: WorkSheet): {
  walkInDoors: PricingLookup;
  windows: PricingLookup;
} {
  const data = sheetToArray(ws);

  const walkInDoors: PricingLookup = {};
  const windows: PricingLookup = {};

  // Sheet layout: walk-in doors (rows 0-~12), then a "Price/Qty/Total" header,
  // then windows (rows ~18+), then another "Price/Qty/Total" header.
  // Detect the divider row to switch sections.
  let section: "doors" | "windows" = "doors";

  for (let r = 0; r < Math.min(40, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    const name = cleanHeader(row[0]);

    // "Price/Qty/Total" header row marks end of a section
    const hasCalcHeader = row.some(
      (v: unknown) => typeof v === "string" && v.trim().toLowerCase() === "price"
    );
    if (hasCalcHeader) {
      if (section === "doors") section = "windows";
      continue;
    }

    if (!name) continue;
    const price = num(row[1]);
    if (price <= 0) continue;
    if (name.toLowerCase().includes("total") || name.toLowerCase().includes("qty")) continue;

    if (section === "doors") {
      walkInDoors[name] = price;
    } else {
      windows[name] = price;
    }
  }

  return { walkInDoors, windows };
}

/**
 * Parse widespan "Roll Up Door" sheet.
 * Layout: Column A = size ("6x6", "6x7", ..., "12x16")
 *         Column B = base price
 *         Column C = price with header/install
 *         Column F = header flag (0 or 1)
 *   Header pricing at top: "10-15' Header" = 515, "16'-20' Header" = 580
 *
 * Returns { rollUpDoors, headerPrices }
 */
export function readWidespanRollUpDoors(ws: WorkSheet): {
  rollUpDoors: PricingLookup;
  rollUpDoorsWithHeader: PricingLookup;
  headerSmall: number;
  headerLarge: number;
} {
  const data = sheetToArray(ws);

  const rollUpDoors: PricingLookup = {};
  const rollUpDoorsWithHeader: PricingLookup = {};
  let headerSmall = 515;
  let headerLarge = 580;

  for (let r = 0; r < Math.min(40, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    const name = cleanHeader(row[0]);

    // Check for header pricing rows
    if (name.includes("Header") || name.includes("header")) {
      const price = num(row[1]) || num(row[2]);
      if (name.includes("10") || name.includes("15")) {
        headerSmall = price;
      } else if (name.includes("16") || name.includes("20")) {
        headerLarge = price;
      }
      continue;
    }

    // Roll-up door sizes (NxN format)
    if (name.match(/^\d+x\d+$/)) {
      const basePrice = num(row[1]);
      const withHeaderPrice = num(row[2]);
      if (basePrice > 0) {
        rollUpDoors[name] = basePrice;
        if (withHeaderPrice > 0) {
          rollUpDoorsWithHeader[name] = withHeaderPrice;
        }
      }
    }
  }

  return { rollUpDoors, rollUpDoorsWithHeader, headerSmall, headerLarge };
}
