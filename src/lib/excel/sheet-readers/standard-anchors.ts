import type { WorkSheet } from "xlsx";
import type { PricingLookup } from "@/types/pricing";
import { sheetToArray, num, cleanHeader } from "./utils";

/**
 * Parse "Pricing - Anchors" sheet.
 * Layout: Column A = "{width}x{endCount}" keys, Column B = anchor counts
 *   e.g. "12x0"=2, "12x1"=4, "12x2"=6
 *
 * Note: Anchors were "removed from" pricing in some versions,
 * meaning they may be included in base price. We still parse
 * the counts for reference.
 *
 * Returns lookup[widthXendCount] → anchorCount
 */
export function readAnchors(ws: WorkSheet): PricingLookup {
  const data = sheetToArray(ws);
  const anchors: PricingLookup = {};

  for (let r = 0; r < Math.min(40, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    const key = cleanHeader(row[0]);
    if (!key || !key.includes("x")) continue;
    anchors[key] = num(row[1]);
  }

  return anchors;
}
