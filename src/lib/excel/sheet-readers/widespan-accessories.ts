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

  let section: "doors" | "windows" | "unknown" = "unknown";

  for (let r = 0; r < Math.min(40, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    const name = cleanHeader(row[0]);
    if (!name) continue;

    // Detect section switches
    if (name.toLowerCase().includes("door") && num(row[1]) > 0) {
      section = "doors";
    } else if (name.toLowerCase().includes("window") && num(row[1]) > 0) {
      section = "windows";
    } else if (name.toLowerCase().includes("window") && !num(row[1])) {
      section = "windows";
      continue;
    }

    const price = num(row[1]);
    if (price <= 0) continue;

    // Skip calculation rows
    if (name.toLowerCase().includes("total") || name.toLowerCase().includes("qty")) continue;

    if (section === "doors" || name.toLowerCase().includes("door") || name.toLowerCase().includes("frame out")) {
      walkInDoors[name] = price;
    } else if (section === "windows" || name.toLowerCase().includes("window") || name.toLowerCase().includes("pane")) {
      windows[name] = price;
    }
  }

  // If Frame Out appears in doors, also add to windows
  // (both sections have a Frame Out option)
  if (!windows["Frame Out"] && walkInDoors["Frame Out"]) {
    // Scan for window frame out separately
    for (let r = 15; r < Math.min(40, data.length); r++) {
      const row = data[r];
      if (!row) continue;
      const name = cleanHeader(row[0]);
      if (name === "Frame Out") {
        windows["Frame Out"] = num(row[1]);
        break;
      }
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
