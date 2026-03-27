import * as XLSX from "xlsx";

export interface SpreadsheetDetection {
  type: "standard" | "widespan";
  states: string[];
  sheetCount: number;
  filenameStates: string[];
}

/**
 * Auto-detect whether a workbook is standard or widespan.
 *
 * Heuristics:
 * - Widespan has "Changers" (not "Pricing - Changers"), "Base Price" (not "Pricing - Base")
 * - Standard has 21 sheets, widespan has 14
 * - Extract states from title cell or Plans & Calcs sheet
 */
export function detectSpreadsheetType(workbook: XLSX.WorkBook, filename?: string): SpreadsheetDetection {
  const sheetNames = workbook.SheetNames;
  const sheetCount = sheetNames.length;

  // Widespan has plain "Changers" sheet, standard has "Pricing - Changers"
  const hasPlainChangers = sheetNames.includes("Changers");
  const hasPricingChangers = sheetNames.includes("Pricing - Changers");

  let type: "standard" | "widespan";
  if (hasPlainChangers && !hasPricingChangers) {
    type = "widespan";
  } else if (hasPricingChangers) {
    type = "standard";
  } else {
    // Fallback to sheet count
    type = sheetCount <= 16 ? "widespan" : "standard";
  }

  // Extract states from the first sheet title
  const states = extractStates(workbook);

  // Extract states from the filename (most reliable source)
  const filenameStates = filename ? extractStatesFromFilename(filename) : [];

  return { type, states: states.length > 0 ? states : filenameStates, sheetCount, filenameStates };
}

function extractStates(workbook: XLSX.WorkBook): string[] {
  // Try to read states from the Quote Sheet title or first sheet
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!firstSheet) return [];

  // Look for a cell that contains state abbreviations
  // The title cell typically contains something like "AZ CO UT 1 5 26"
  const titleCell = firstSheet["A1"] || firstSheet["B1"] || firstSheet["A2"];
  if (!titleCell?.v) return [];

  const title = String(titleCell.v);
  const statePattern = /\b([A-Z]{2})\b/g;
  const matches = title.match(statePattern) || [];

  // Filter to valid US state abbreviations
  const validStates = new Set([
    "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN",
    "IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV",
    "NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN",
    "TX","UT","VT","VA","WA","WV","WI","WY",
  ]);

  return matches.filter((s) => validStates.has(s));
}

const VALID_STATES = new Set([
  "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN",
  "IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV",
  "NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN",
  "TX","UT","VT","VA","WA","WV","WI","WY",
]);

/**
 * Extract state abbreviations from the filename.
 * Handles patterns like:
 *   "TX-Dealer-Calcs-3-5-26-1.xlsx"  →  ["TX"]
 *   "AZ CO UT 1 5 26.xlsx"           →  ["AZ","CO","UT"]
 *   "MI-PA-Wide-Span-3-5-26-2.xlsx"  →  ["MI","PA"]
 *   "CA-Dealer-Calcs-3-5-26-1.xlsx"  →  ["CA"]
 */
function extractStatesFromFilename(filename: string): string[] {
  // Remove extension
  const name = filename.replace(/\.[^.]+$/, "");

  // Strategy 1: Leading state codes separated by hyphens or spaces
  // e.g. "TX-Dealer-...", "AZ CO UT 1 5 26", "MI-PA-Wide-Span-..."
  const states: string[] = [];
  // Split on hyphens and spaces, collect leading valid state codes
  const parts = name.split(/[-\s]+/);
  for (const part of parts) {
    if (VALID_STATES.has(part)) {
      states.push(part);
    } else {
      // Stop at first non-state token
      break;
    }
  }

  return states;
}
