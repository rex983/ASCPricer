import type { WorkSheet } from "xlsx";
import type { PricingMatrix, PricingLookup } from "@/types/pricing";
import { sheetToArray, num, cleanHeader, readMatrix } from "./utils";

/**
 * Parse "Snow - Truss Spacing" sheet.
 * Headers: Engineering codes like "E-105-12-STD", "O-105-18-AFV"
 * Row keys: Snow load codes like "T-30GL", "M-40GL"
 *
 * Returns matrix[snowCode][configCode] → spacing (inches)
 * (transposed so we look up by snowCode first, then configKey)
 */
export function readTrussSpacing(ws: WorkSheet): PricingMatrix {
  const data = sheetToArray(ws);
  return readMatrix(data, {
    headerRow: 0,
    dataStartRow: 1,
    rowKeyCol: 0,
    dataStartCol: 1,
    transpose: true,
  });
}

/**
 * Parse "Snow - Trusses" sheet.
 * Headers are "{width}-{state}" like "12-OH", "24-AZ", "18-MI"
 * Rows are building lengths (20, 25, 30, ...).
 *
 * Returns matrix["{width}-{state}"]["{length}"] → count
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
 * Left side: Hat channel spacing. Row keys are "{bucketedTrussSpacing}-{snowCode}"
 *   e.g. "60-T-20LL", "48-S-30GL". Column headers are wind speeds (105,115,...,180).
 * Right side: Original hat channel counts by state/width headers.
 *
 * Returns { spacing, originalCounts }
 */
export function readHatChannels(ws: WorkSheet): {
  spacing: PricingMatrix;
  originalCounts: PricingMatrix;
} {
  const data = sheetToArray(ws);
  const headers = data[0] || [];

  const spacing: PricingMatrix = {};
  const originalCounts: PricingMatrix = {};

  // Find the "Original Hat Channels" section by looking for the keyword "Original"
  // in headers, or for a column with state codes (CA, AZ, OH, etc.) as data
  let rightSectionStart = -1;
  for (let c = 8; c < headers.length; c++) {
    const h = cleanHeader(headers[c]);
    if (h.toLowerCase().includes("original")) {
      rightSectionStart = c;
      break;
    }
  }
  // Fallback: find first column after wind speeds that has state codes in its data
  if (rightSectionStart < 0) {
    for (let c = 8; c < headers.length; c++) {
      const h = cleanHeader(headers[c]);
      // Skip wind speed columns and blanks
      if (["105", "115", "130", "140", "155", "165", "180", "", "0"].includes(h)) continue;
      // Check if data rows have state codes
      let stateCount = 0;
      for (let r = 1; r < Math.min(10, data.length); r++) {
        const v = cleanHeader(data[r]?.[c] ?? "");
        if (v.match(/^[A-Z]{2}$/)) stateCount++;
      }
      if (stateCount >= 3) {
        rightSectionStart = c;
        break;
      }
    }
  }

  // Read spacing (left side) — transpose so spacing[rowKey][windSpeed] → value
  const spacingEndCol = rightSectionStart > 0 ? rightSectionStart : 8;
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row) break;
    const rowKey = cleanHeader(row[0]);
    if (!rowKey) break;

    if (!spacing[rowKey]) spacing[rowKey] = {};
    for (let c = 1; c < spacingEndCol; c++) {
      const colKey = cleanHeader(headers[c]);
      if (!colKey || colKey === "0") continue;
      spacing[rowKey][colKey] = num(row[c]);
    }
  }

  // Read original counts (right side)
  // Layout: first col of right section = state codes (or "Original Hat Channels" label)
  // Header row has width values (12, 18, 20, 22, 24, 26, 28, 30)
  // Data rows: state code in first col, then counts per width
  if (rightSectionStart > 0) {
    // Find the column with state codes and the width header columns
    // The header row for widths may be row 0 or a different row
    let stateCol = rightSectionStart;
    let widthStartCol = rightSectionStart + 1;

    // Check if rightSectionStart is the "Original" label column
    const firstHeader = cleanHeader(headers[rightSectionStart]);
    if (firstHeader.toLowerCase().includes("original") || !firstHeader.match(/^\d+$/)) {
      stateCol = rightSectionStart;
      widthStartCol = rightSectionStart + 1;
    }

    // Find the header row for width values (could be row 0 or the same row)
    // Width headers should be numbers like 12, 18, 20, etc.
    let widthHeaders: string[] = [];
    for (let c = widthStartCol; c < headers.length; c++) {
      const h = cleanHeader(headers[c]);
      if (h && num(h) >= 12 && num(h) <= 30) {
        widthHeaders.push(h);
      } else {
        widthHeaders.push("");
      }
    }

    // Parse data rows: state code → { width → count }
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      if (!row) break;
      const state = cleanHeader(row[stateCol]);
      if (!state || !state.match(/^[A-Z]{2}$/)) continue;

      if (!originalCounts[state]) originalCounts[state] = {};
      for (let c = widthStartCol; c < row.length; c++) {
        const widthKey = widthHeaders[c - widthStartCol];
        if (!widthKey) continue;
        const count = num(row[c]);
        if (count > 0) {
          originalCounts[state][widthKey] = count;
        }
      }
    }
  }

  return { spacing, originalCounts };
}

/**
 * Parse "Snow - Girts" sheet.
 * Two sections:
 * Left: spacing matrix — rows are truss spacing buckets (60/54/48/42/36),
 *   cols are wind speeds (105-180).
 *   Returns girtSpacing["{bucketedTrussSpacing}"]["{windSpeed}"] → spacing
 * Right: original girt counts by leg height.
 *   Returns girtCountsByHeight[height] → count
 */
export function readGirtSpacing(ws: WorkSheet): {
  spacing: PricingMatrix;
  girtCountsByHeight: PricingLookup;
} {
  const data = sheetToArray(ws);
  const headers = data[0] || [];

  const spacing: PricingMatrix = {};
  const girtCountsByHeight: PricingLookup = {};

  // Find the split between left (spacing) and right (original counts) sections
  let rightSectionStart = -1;
  for (let c = 1; c < headers.length; c++) {
    const h = cleanHeader(headers[c]);
    if (h === "" || h === "0") {
      if (rightSectionStart < 0) {
        const next = cleanHeader(headers[c + 1]);
        if (next && !["105", "115", "130", "140", "155", "165", "180"].includes(next)) {
          rightSectionStart = c + 1;
          break;
        }
      }
    }
  }

  // Left side: spacing matrix — girtSpacing[trussSpacing][windSpeed] → spacing
  const spacingEndCol = rightSectionStart > 0 ? rightSectionStart : headers.length;
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row) break;
    const rowKey = cleanHeader(row[0]);
    if (!rowKey) break;

    if (!spacing[rowKey]) spacing[rowKey] = {};
    for (let c = 1; c < spacingEndCol; c++) {
      const colKey = cleanHeader(headers[c]);
      if (!colKey || colKey === "0") continue;
      spacing[rowKey][colKey] = num(row[c]);
    }
  }

  // Right side: original girt counts by height
  // Format: rows of height → count pairs
  if (rightSectionStart > 0) {
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      if (!row) break;
      const heightKey = cleanHeader(row[rightSectionStart]);
      const count = num(row[rightSectionStart + 1]);
      if (!heightKey) continue;
      // Heights like "0-9", "10-12", "13-15", "16-20" or individual numbers
      if (count > 0) {
        girtCountsByHeight[heightKey] = count;
      }
    }
  }

  return { spacing, girtCountsByHeight };
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

  // Main spacing matrix: verticalSpacing[windSpeed][height] → spacing
  const spacing = readMatrix(data, {
    headerRow: 0,
    dataStartRow: 1,
    rowKeyCol: 0,
    dataStartCol: 1,
  });

  // Original vertical counts (around row 8-13 depending on number of wind speeds)
  const originalCounts: PricingLookup = {};
  for (let r = 7; r < Math.min(20, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    const label = cleanHeader(row[0]);
    if (label.toLowerCase().includes("original")) {
      // This row has width values (12, 18, 20, 22, 24, 26, 28, 30)
      // Next row has the corresponding counts
      const widthRow = row;
      const countRow = data[r + 1];
      if (countRow) {
        for (let c = 1; c < widthRow.length; c++) {
          const width = cleanHeader(widthRow[c]);
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
 *
 * Actual spreadsheet layout:
 *   Rows 0-5: Wind load bucketing (input MPH → standard category)
 *     Row 3: input MPH values across columns (90, 100, 120, 130, 140, 150, 165)
 *     Row 5: bucketed output (105, 115, 130, 140, 155, 165, 180)
 *   Rows 8-9: Snow load label → code mapping
 *   Rows 25-27: Height classification + Feet Used
 *     Row 25 (or first found): height values (1-20)
 *     Row 26: S/M/T prefix per height
 *     Row 27: Feet Used per height (for truss leg surcharge)
 *   Rows 57-69: Per-state pricing section
 *     Row 58: state codes across columns (WI, OH, TX, NM, AZ, PA, CA, ...)
 *     Rows 59-66: truss prices — col 0 = width label, remaining cols = price per state
 *     Row 67: Pie Truss Price ($15/ft per state)
 *     Row 68: Channel Price Per Ft ($2-$2.50 per state)
 *     Row 69: Tubing Price Per Ft ($3-$4 per state)
 */
export function readSnowChangers(ws: WorkSheet): {
  windLoadBuckets: PricingLookup;
  snowLoadOptions: PricingLookup;
  heightClassification: PricingLookup;
  feetUsedByHeight: PricingLookup;
  pieTrussPrice: PricingLookup;
  trussPriceByWidthState: PricingMatrix;
  channelPriceByState: PricingLookup;
  tubingPriceByState: PricingLookup;
} {
  const data = sheetToArray(ws);

  const windLoadBuckets: PricingLookup = {};
  const snowLoadOptions: PricingLookup = {};
  const heightClassification: PricingLookup = {};
  const feetUsedByHeight: PricingLookup = {};
  const pieTrussPrice: PricingLookup = {};
  const trussPriceByWidthState: PricingMatrix = {};
  const channelPriceByState: PricingLookup = {};
  const tubingPriceByState: PricingLookup = {};

  // ── Wind load buckets (rows 0-1) ──
  // Row 0: "Wind Load" label in col 0, then sequential MPH values (0, 1, 2, ..., 180) in cols 1+
  // Row 1: "Roof" label in col 0, then bucketed output per column (105, 105, ..., 115, ..., 130, ...)
  // The column index maps to input MPH: col N has input = N-1, output = row1[N]
  if (data[0] && data[1]) {
    for (let c = 1; c < data[0].length; c++) {
      const inputMph = num(data[0][c]);
      const bucketMph = num(data[1][c]);
      if (inputMph >= 80 && inputMph <= 200 && bucketMph >= 100 && bucketMph <= 200) {
        windLoadBuckets[String(inputMph)] = bucketMph;
      }
    }
  }

  // ── Snow load option labels (rows 5-12) ──
  for (let r = 4; r < Math.min(14, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const val = cleanHeader(row[c]);
      if (val.match(/^\d+[GL]L$/i)) {
        const desc = c > 0 ? cleanHeader(row[c - 1]) : "";
        if (desc) snowLoadOptions[desc] = num(val.replace(/[GL]L$/i, ""));
      }
      if (val.toLowerCase().includes("roof load") || val.toLowerCase().includes("ground load")) {
        const nextVal = cleanHeader(row[c + 1] ?? "");
        if (nextVal.match(/^\d+[GL]L$/i)) {
          snowLoadOptions[val] = num(nextVal.replace(/[GL]L$/i, ""));
        }
      }
    }
  }

  // ── Height classification + Feet Used (rows 20-30) ──
  // Find row with S/M/T values, then heights row above and feetUsed row below
  for (let r = 15; r < Math.min(35, data.length); r++) {
    const row = data[r];
    if (!row) continue;

    // Count S/M/T values in this row
    let smtCount = 0;
    for (let c = 0; c < row.length; c++) {
      const v = cleanHeader(row[c]);
      if (v === "S" || v === "M" || v === "T") smtCount++;
    }

    if (smtCount >= 5) {
      // This is the S/M/T row — heights should be in the row above
      const heightRow = data[r - 1];
      const feetRow = data[r + 1];
      if (heightRow) {
        for (let c = 0; c < row.length; c++) {
          const prefix = cleanHeader(row[c]);
          if (prefix !== "S" && prefix !== "M" && prefix !== "T") continue;
          const h = num(heightRow[c]);
          if (h >= 1 && h <= 30) {
            heightClassification[String(h)] = prefix === "S" ? 0 : prefix === "M" ? 1 : 2;
            if (feetRow) {
              const feet = num(feetRow[c]);
              feetUsedByHeight[String(h)] = feet;
            }
          }
        }
      }
      break;
    }
  }

  // ── Per-state pricing (rows 55-75) ──
  // Find the state code row (row with multiple 2-letter codes)
  let stateCodeRow = -1;
  let stateCodes: string[] = [];
  for (let r = 50; r < Math.min(75, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    let codeCount = 0;
    const codes: string[] = [];
    for (let c = 1; c < row.length; c++) {
      const v = cleanHeader(row[c]);
      if (v.match(/^[A-Z]{2}$/)) {
        codeCount++;
        codes.push(v);
      } else {
        codes.push("");
      }
    }
    if (codeCount >= 5) {
      stateCodeRow = r;
      stateCodes = codes;
      break;
    }
  }

  if (stateCodeRow >= 0) {
    // Parse truss price rows (rows after state codes, labeled with width)
    for (let r = stateCodeRow + 1; r < Math.min(stateCodeRow + 12, data.length); r++) {
      const row = data[r];
      if (!row) continue;
      const label = cleanHeader(row[0]).toLowerCase();

      if (label.includes("pie")) {
        // Pie truss price row — must check BEFORE "truss" since label contains both words
        for (let c = 1; c < row.length && c - 1 < stateCodes.length; c++) {
          const st = stateCodes[c - 1];
          if (!st) continue;
          const price = num(row[c]);
          if (price > 0 && price < 100) {
            pieTrussPrice[st] = price;
          }
        }
      } else if (label.includes("wide") || label.includes("truss")) {
        // Extract width number from label like "12' Wide Truss" or "12"
        const widthMatch = label.match(/(\d+)/);
        if (widthMatch) {
          const w = widthMatch[1];
          for (let c = 1; c < row.length && c - 1 < stateCodes.length; c++) {
            const st = stateCodes[c - 1];
            if (!st) continue;
            const price = num(row[c]);
            if (price > 50 && price < 1000) {
              if (!trussPriceByWidthState[st]) trussPriceByWidthState[st] = {};
              trussPriceByWidthState[st][w] = price;
            }
          }
        }
      } else if (label.includes("channel")) {
        for (let c = 1; c < row.length && c - 1 < stateCodes.length; c++) {
          const st = stateCodes[c - 1];
          if (!st) continue;
          const price = num(row[c]);
          if (price >= 1 && price <= 10) {
            channelPriceByState[st] = price;
          }
        }
      } else if (label.includes("tubing") || label.includes("tube")) {
        for (let c = 1; c < row.length && c - 1 < stateCodes.length; c++) {
          const st = stateCodes[c - 1];
          if (!st) continue;
          const price = num(row[c]);
          if (price >= 1 && price <= 10) {
            tubingPriceByState[st] = price;
          }
        }
      }
    }
  }

  return {
    windLoadBuckets,
    snowLoadOptions,
    heightClassification,
    feetUsedByHeight,
    pieTrussPrice,
    trussPriceByWidthState,
    channelPriceByState,
    tubingPriceByState,
  };
}
