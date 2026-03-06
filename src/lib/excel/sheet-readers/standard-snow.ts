import type { WorkSheet } from "xlsx";
import type { PricingMatrix, PricingLookup } from "@/types/pricing";
import { sheetToArray, num, cleanHeader, readMatrix } from "./utils";

/**
 * Parse "Snow - Truss Spacing" sheet.
 * Headers: Engineering codes like "E-105-12-STD", "O-105-18-AFV"
 * Row keys: Snow load codes like "T-30GL", "M-40GL"
 *
 * Returns matrix[configCode][snowLoad] → spacing (inches)
 */
export function readTrussSpacing(ws: WorkSheet): PricingMatrix {
  const data = sheetToArray(ws);
  return readMatrix(data, {
    headerRow: 0,
    dataStartRow: 1,
    rowKeyCol: 0,
    dataStartCol: 1,
  });
}

/**
 * Parse "Snow - Trusses" sheet.
 * Contains original truss counts by state/region and width.
 * Headers include state codes like "12-OH", "18-OH", "12-MI", etc.
 *
 * Returns matrix[stateWidth][config] → count
 */
export function readTrussCounts(ws: WorkSheet): PricingMatrix {
  const data = sheetToArray(ws);
  return readMatrix(data, {
    headerRow: 0,
    dataStartRow: 1,
    rowKeyCol: 0,
    dataStartCol: 1,
  });
}

/**
 * Parse "Snow - Hat Channels" sheet.
 * Left side: Hat channel spacing by trussSpacing-snowLoad key and wind speed.
 * Right side: Original hat channel counts by state and width.
 *
 * Returns { spacing, originalCounts }
 */
export function readHatChannels(ws: WorkSheet): {
  spacing: PricingMatrix;
  originalCounts: PricingMatrix;
} {
  const data = sheetToArray(ws);
  const headers = data[0] || [];

  // Find the split between spacing and original counts sections
  // Spacing section has wind speed headers (105, 115, 130, ...)
  // Original counts section has state/width headers

  const spacing: PricingMatrix = {};
  const originalCounts: PricingMatrix = {};

  // Left side: spacing matrix
  // Headers are wind speeds: 105, 115, 130, 140, 155, 165, 180
  let rightSectionStart = -1;
  for (let c = 1; c < headers.length; c++) {
    const h = cleanHeader(headers[c]);
    if (h === "" || h === "0") {
      // Check for section break
      if (rightSectionStart < 0) {
        const next = cleanHeader(headers[c + 1]);
        if (next && !["105", "115", "130", "140", "155", "165", "180"].includes(next)) {
          rightSectionStart = c + 1;
        }
      }
    }
  }

  // Read spacing (left side)
  const spacingEndCol = rightSectionStart > 0 ? rightSectionStart : 8;
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row) break;
    const rowKey = cleanHeader(row[0]);
    if (!rowKey) break;

    for (let c = 1; c < spacingEndCol; c++) {
      const colKey = cleanHeader(headers[c]);
      if (!colKey || colKey === "0") continue;
      if (!spacing[colKey]) spacing[colKey] = {};
      spacing[colKey][rowKey] = num(row[c]);
    }
  }

  // Read original counts (right side)
  if (rightSectionStart > 0) {
    const rightHeaders = data[rightSectionStart > 0 ? 0 : 0] || [];
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      if (!row) break;
      // Find row key in the right section
      for (let c = rightSectionStart; c < headers.length; c++) {
        const colKey = cleanHeader(rightHeaders[c]);
        if (!colKey || colKey === "0") continue;
        const rowKey = cleanHeader(row[rightSectionStart - 1]) || cleanHeader(row[0]);
        if (!rowKey) break;
        if (!originalCounts[colKey]) originalCounts[colKey] = {};
        originalCounts[colKey][rowKey] = num(row[c]);
      }
    }
  }

  return { spacing, originalCounts };
}

/**
 * Parse "Snow - Girts" sheet.
 * Similar structure to Hat Channels: girt spacing by wind and config.
 */
export function readGirtSpacing(ws: WorkSheet): PricingMatrix {
  const data = sheetToArray(ws);
  return readMatrix(data, {
    headerRow: 0,
    dataStartRow: 1,
    rowKeyCol: 0,
    dataStartCol: 1,
  });
}

/**
 * Parse "Snow - Verticals" sheet.
 * Vertical spacing by wind speed (rows) and leg height (columns).
 * Also contains original vertical counts by width.
 */
export function readVerticals(ws: WorkSheet): {
  spacing: PricingMatrix;
  originalCounts: PricingLookup;
} {
  const data = sheetToArray(ws);
  const headers = data[0] || [];

  // Main spacing matrix
  const spacing = readMatrix(data, {
    headerRow: 0,
    dataStartRow: 1,
    rowKeyCol: 0,
    dataStartCol: 1,
  });

  // Original vertical counts (usually around row 12-13)
  const originalCounts: PricingLookup = {};
  for (let r = 10; r < Math.min(20, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    const label = cleanHeader(row[0]);
    if (label.toLowerCase().includes("original")) {
      // Next row typically has width→count mapping
      const countRow = data[r + 1];
      if (countRow) {
        for (let c = 1; c < headers.length; c++) {
          const width = cleanHeader(headers[c]);
          if (width && num(width) >= 12) {
            originalCounts[width] = num(countRow[c]);
          }
        }
      }
      break;
    }
  }

  return { spacing, originalCounts };
}

/**
 * Parse "Snow - Diagonal Bracing" sheet.
 * Contains state-by-state wind thresholds for when DB is required,
 * and pricing calculations.
 */
export function readDiagonalBracing(ws: WorkSheet): {
  windThresholdByState: PricingLookup;
  baseBracePrice: number;
  tallSurcharge: number;
} {
  const data = sheetToArray(ws);

  const windThresholdByState: PricingLookup = {};
  let baseBracePrice = 90;
  let tallSurcharge = 50;

  // Scan for state thresholds (typically around rows 12-13)
  for (let r = 0; r < Math.min(30, data.length); r++) {
    const row = data[r];
    if (!row) continue;

    // Look for state abbreviations with wind thresholds
    for (let c = 0; c < row.length - 1; c++) {
      const cellStr = cleanHeader(row[c]);
      // State codes are 2 letter uppercase
      if (cellStr.match(/^[A-Z]{2}$/) && num(row[c + 1]) >= 100) {
        windThresholdByState[cellStr] = num(row[c + 1]);
      }
    }

    // Look for brace pricing
    const label = cleanHeader(row[0]);
    if (label.toLowerCase().includes("price") || label.toLowerCase().includes("cost")) {
      for (let c = 1; c < row.length; c++) {
        const val = num(row[c]);
        if (val >= 50 && val <= 200) {
          if (!baseBracePrice || val < baseBracePrice) baseBracePrice = val;
          break;
        }
      }
    }
  }

  return { windThresholdByState, baseBracePrice, tallSurcharge };
}

/**
 * Parse "Snow - Changers" sheet.
 * Resolves wind load, snow load, and other engineering parameters.
 *
 * Sections:
 *   R0-R5: Wind Load mapping (input MPH → actual wind category)
 *   R8-R13: Snow Load mapping (description → code like "20LL", "30GL")
 *   R16-R21: Hat Channel chart selection
 *   R25-R29: Height → S/M/T classification and feet
 */
export function readSnowChangers(ws: WorkSheet): {
  windLoadMapping: PricingLookup;
  snowLoadMapping: PricingLookup;
} {
  const data = sheetToArray(ws);

  const windLoadMapping: PricingLookup = {};
  const snowLoadMapping: PricingLookup = {};

  // Wind load section (rows 0-5)
  for (let r = 0; r < Math.min(8, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 0; c < row.length - 1; c++) {
      const key = cleanHeader(row[c]);
      const val = num(row[c + 1]);
      if (key && val >= 90 && val <= 200) {
        windLoadMapping[key] = val;
      }
    }
  }

  // Snow load section (rows 8-13)
  for (let r = 8; r < Math.min(16, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const val = cleanHeader(row[c]);
      if (val.match(/^\d+[GL]L$/)) {
        // Found a snow code, map description to code
        const desc = cleanHeader(row[c - 1]) || cleanHeader(row[0]);
        if (desc) snowLoadMapping[desc] = num(val.replace(/[GL]L$/, ""));
      }
    }
  }

  return { windLoadMapping, snowLoadMapping };
}
