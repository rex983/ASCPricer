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

  // Find the split between spacing and original counts sections
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
  if (rightSectionStart > 0) {
    const rightHeaders = data[0] || [];
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      if (!row) break;
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
 *   R0-R1: Wind Load Buckets — input MPH → category (105/115/130/140/155/165/180)
 *   R5-R6: Snow Load option labels — "30 Ground Load"→"30GL", "20 Roof Load"→"20LL"
 *   R15-R16: Height classification — height → S/M/T prefix
 *   R36-R46: Per-state pricing:
 *     Truss price by width+state (rows 36-43)
 *     Channel price by state (row 45)
 *     Tubing price by state (row 46)
 */
export function readSnowChangers(ws: WorkSheet): {
  windLoadBuckets: PricingLookup;
  snowLoadOptions: PricingLookup;
  heightClassification: PricingLookup;
  trussPriceByWidthState: PricingMatrix;
  channelPriceByState: PricingLookup;
  tubingPriceByState: PricingLookup;
} {
  const data = sheetToArray(ws);

  const windLoadBuckets: PricingLookup = {};
  const snowLoadOptions: PricingLookup = {};
  const heightClassification: PricingLookup = {};
  const trussPriceByWidthState: PricingMatrix = {};
  const channelPriceByState: PricingLookup = {};
  const tubingPriceByState: PricingLookup = {};

  // Wind load buckets (rows 0-3): input MPH → bucket category
  for (let r = 0; r < Math.min(5, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 0; c < row.length - 1; c += 2) {
      const inputMph = num(row[c]);
      const bucketMph = num(row[c + 1]);
      if (inputMph >= 85 && inputMph <= 200 && bucketMph >= 100 && bucketMph <= 200) {
        windLoadBuckets[String(inputMph)] = bucketMph;
      }
    }
  }

  // Snow load option labels (rows 5-10)
  for (let r = 4; r < Math.min(14, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const val = cleanHeader(row[c]);
      // Match snow codes like "20LL", "30GL", "40GL"
      if (val.match(/^\d+[GL]L$/i)) {
        // Map description → code
        const desc = cleanHeader(row[c - 1]) || cleanHeader(row[0]);
        if (desc) snowLoadOptions[desc] = num(val.replace(/[GL]L$/i, ""));
      }
      // Also capture the raw label mapping (e.g. "20 Roof Load" → "20LL")
      if (val.toLowerCase().includes("roof load") || val.toLowerCase().includes("ground load")) {
        const nextVal = cleanHeader(row[c + 1]);
        if (nextVal.match(/^\d+[GL]L$/i)) {
          snowLoadOptions[val] = num(nextVal.replace(/[GL]L$/i, ""));
        }
      }
    }
  }

  // Height classification (rows 15-20): height → S/M/T prefix
  for (let r = 14; r < Math.min(25, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 0; c < row.length - 1; c += 2) {
      const heightVal = cleanHeader(row[c]);
      const prefix = cleanHeader(row[c + 1]);
      if (prefix.match(/^[SMT]$/) && heightVal) {
        heightClassification[heightVal] = prefix === "S" ? 0 : prefix === "M" ? 1 : 2;
      }
    }
    // Also scan for "S", "M", "T" in cells with height ranges
    for (let c = 0; c < row.length; c++) {
      const val = cleanHeader(row[c]);
      if (val === "S" || val === "M" || val === "T") {
        // Check neighboring cells for height values
        const prev = cleanHeader(row[c - 1]);
        if (prev && num(prev) >= 6 && num(prev) <= 20) {
          heightClassification[prev] = val === "S" ? 0 : val === "M" ? 1 : 2;
        }
      }
    }
  }

  // Per-state truss pricing (rows 36-43)
  // Format: state in col 0, then width ranges with prices
  for (let r = 34; r < Math.min(48, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    const stateOrLabel = cleanHeader(row[0]);

    // Detect state rows (2-letter state codes)
    if (stateOrLabel.match(/^[A-Z]{2}$/)) {
      // Parse width-price pairs from this row
      for (let c = 1; c < row.length - 1; c += 2) {
        const widthRange = cleanHeader(row[c]);
        const price = num(row[c + 1]);
        if (price > 0 && widthRange) {
          // widthRange might be "12-24" or "26-30"
          if (!trussPriceByWidthState[stateOrLabel]) trussPriceByWidthState[stateOrLabel] = {};
          trussPriceByWidthState[stateOrLabel][widthRange] = price;
        }
      }
    }

    // Channel price by state (look for "channel" label)
    if (stateOrLabel.toLowerCase().includes("channel") || stateOrLabel.toLowerCase().includes("hat")) {
      for (let c = 1; c < row.length - 1; c += 2) {
        const state = cleanHeader(row[c]);
        const price = num(row[c + 1]);
        if (state.match(/^[A-Z]{2}$/) && price > 0 && price < 10) {
          channelPriceByState[state] = price;
        }
      }
    }

    // Tubing price by state (look for "tubing" label)
    if (stateOrLabel.toLowerCase().includes("tubing") || stateOrLabel.toLowerCase().includes("tube")) {
      for (let c = 1; c < row.length - 1; c += 2) {
        const state = cleanHeader(row[c]);
        const price = num(row[c + 1]);
        if (state.match(/^[A-Z]{2}$/) && price > 0 && price < 10) {
          tubingPriceByState[state] = price;
        }
      }
    }
  }

  // Fallback: scan more broadly for channel/tubing pricing
  if (Object.keys(channelPriceByState).length === 0) {
    for (let r = 30; r < Math.min(50, data.length); r++) {
      const row = data[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const val = cleanHeader(row[c]);
        if (val.match(/^[A-Z]{2}$/)) {
          // Check if there's a price in adjacent cells
          const nextVal = num(row[c + 1]);
          if (nextVal >= 1.5 && nextVal <= 5) {
            // Could be channel or tubing price
            const labelAbove = r > 0 ? cleanHeader(data[r - 1]?.[c] ?? "") : "";
            if (labelAbove.toLowerCase().includes("channel")) {
              channelPriceByState[val] = nextVal;
            } else if (labelAbove.toLowerCase().includes("tub")) {
              tubingPriceByState[val] = nextVal;
            }
          }
        }
      }
    }
  }

  return {
    windLoadBuckets,
    snowLoadOptions,
    heightClassification,
    trussPriceByWidthState,
    channelPriceByState,
    tubingPriceByState,
  };
}
