import type { WorkSheet } from "xlsx";
import type { PricingMatrix, PricingLookup } from "@/types/pricing";
import { sheetToArray, num, cleanHeader, readMatrix } from "./utils";

/**
 * Parse widespan "Snow Load" sheet.
 * Multi-section sheet containing:
 *
 * 1. Purlin spacing matrix (left): rows are snow codes (S-30GL, G-50GL, etc.),
 *    columns are config codes (E-105-S, O-120-G, etc.) → spacing in inches
 *
 * 2. Truss table (cols T-W): length → truss count and max spacing
 *
 * 3. Verticals table (cols Z-AC): width → vertical count and spacing
 *
 * 4. Purlin required spacing (cols AF-AG): snow code → required spacing
 *
 * 5. Original purlin count (cols AJ-AK): width → count
 *
 * 6. Girt spacing (cols AP-AQ): wind speed → girt spacing
 *
 * 7. Original girts (cols AT-AU): leg height → girt count
 */
export function readWidespanSnowLoad(ws: WorkSheet): {
  purlinSpacing: PricingMatrix;
  trussCountByLength: PricingLookup;
  trussSpacingByLength: PricingLookup;
  verticalCountByWidth: PricingLookup;
  verticalSpacingByWidth: PricingLookup;
  verticalSpacingByWind: PricingLookup;
  purlinRequiredSpacing: PricingLookup;
  originalPurlinByWidth: PricingLookup;
  girtSpacingByWind: PricingLookup;
  originalGirtsByHeight: PricingLookup;
} {
  const data = sheetToArray(ws);
  const headers = data[0] || [];

  // 1. Purlin spacing matrix (first ~18 columns)
  const purlinSpacing: PricingMatrix = {};
  for (let r = 1; r < Math.min(20, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    const rowKey = cleanHeader(row[0]);
    if (!rowKey || !rowKey.match(/^[SG]-/)) continue;

    for (let c = 1; c < Math.min(18, headers.length); c++) {
      const colKey = cleanHeader(headers[c]);
      if (!colKey || !colKey.match(/^[EO]-/)) continue;
      if (!purlinSpacing[colKey]) purlinSpacing[colKey] = {};
      purlinSpacing[colKey][rowKey] = num(row[c]);
    }
  }

  // Scan remaining columns for the supporting lookup tables
  const trussCountByLength: PricingLookup = {};
  const trussSpacingByLength: PricingLookup = {};
  const verticalCountByWidth: PricingLookup = {};
  const verticalSpacingByWidth: PricingLookup = {};
  const verticalSpacingByWind: PricingLookup = {};
  const purlinRequiredSpacing: PricingLookup = {};
  const originalPurlinByWidth: PricingLookup = {};
  const girtSpacingByWind: PricingLookup = {};
  const originalGirtsByHeight: PricingLookup = {};

  // Find column sections by scanning headers for known patterns
  for (let c = 18; c < headers.length; c++) {
    const h = cleanHeader(headers[c]);

    // Truss section: lengths 20-200
    if (h === "Length" || h === "length") {
      // Next columns should have count and spacing
      for (let r = 1; r < Math.min(40, data.length); r++) {
        const row = data[r];
        if (!row) continue;
        const length = num(row[c]);
        if (length >= 20 && length <= 200) {
          const count = num(row[c + 1]);
          const spacing = num(row[c + 2]);
          if (count > 0) trussCountByLength[String(length)] = count;
          if (spacing > 0) trussSpacingByLength[String(length)] = spacing;
        }
      }
    }

    // Verticals section: widths 32-60
    if (h === "Width" || h === "width") {
      for (let r = 1; r < Math.min(20, data.length); r++) {
        const row = data[r];
        if (!row) continue;
        const width = num(row[c]);
        if (width >= 32 && width <= 60) {
          const count = num(row[c + 1]);
          const spacing = num(row[c + 2]);
          if (count > 0) verticalCountByWidth[String(width)] = count;
          if (spacing > 0) verticalSpacingByWidth[String(width)] = spacing;
        }
      }
    }

    // Purlin required spacing: snow code → spacing
    if (h.match(/^[0-9]+[GL]L$/) || h === "0GL") {
      for (let r = 1; r < Math.min(20, data.length); r++) {
        const row = data[r];
        if (!row) continue;
        const code = cleanHeader(row[c]);
        const spacing = num(row[c + 1]);
        if (code.match(/[GL]L$/) && spacing > 0) {
          purlinRequiredSpacing[code] = spacing;
        }
      }
    }

    // Original purlin count by width
    if (h.toLowerCase().includes("original") && h.toLowerCase().includes("purlin")) {
      for (let r = 1; r < Math.min(20, data.length); r++) {
        const row = data[r];
        if (!row) continue;
        const width = num(row[c]);
        const count = num(row[c + 1]);
        if (width >= 32 && width <= 60 && count > 0) {
          originalPurlinByWidth[String(width)] = count;
        }
      }
    }

    // Girt spacing by wind
    if (h.toLowerCase().includes("wind") || h.toLowerCase().includes("girt")) {
      for (let r = 1; r < Math.min(10, data.length); r++) {
        const row = data[r];
        if (!row) continue;
        const wind = num(row[c]);
        const spacing = num(row[c + 1]);
        if (wind >= 90 && wind <= 200 && spacing > 0) {
          girtSpacingByWind[String(wind)] = spacing;
        }
      }
    }

    // Original girts by height
    if (h.toLowerCase().includes("original") && h.toLowerCase().includes("girt")) {
      for (let r = 1; r < Math.min(20, data.length); r++) {
        const row = data[r];
        if (!row) continue;
        const height = num(row[c]);
        const count = num(row[c + 1]);
        if (height >= 8 && height <= 20 && count > 0) {
          originalGirtsByHeight[String(height)] = count;
        }
      }
    }
  }

  // Scan for vertical spacing by wind (90→96", 110→96", 120→80", 130→72")
  // Typically below the verticals-by-width section
  if (Object.keys(verticalSpacingByWind).length === 0) {
    for (let r = 0; r < Math.min(30, data.length); r++) {
      const row = data[r];
      if (!row) continue;
      for (let c = 18; c < row.length - 1; c++) {
        const wind = num(row[c]);
        const spacing = num(row[c + 1]);
        if ([90, 110, 120, 130].includes(wind) && spacing >= 40 && spacing <= 120) {
          // Verify this isn't girt spacing by checking context
          // Vertical spacings tend to be larger (72-96) vs girt (24-60)
          if (spacing >= 60 || Object.keys(verticalSpacingByWind).length > 0) {
            verticalSpacingByWind[String(wind)] = spacing;
          }
        }
      }
    }
  }

  // Fallback: if column-header-based scanning didn't work,
  // scan all cells for the known patterns
  if (Object.keys(purlinRequiredSpacing).length === 0) {
    for (let r = 0; r < Math.min(25, data.length); r++) {
      const row = data[r];
      if (!row) continue;
      for (let c = 20; c < row.length - 1; c++) {
        const key = cleanHeader(row[c]);
        const val = num(row[c + 1]);
        if (key.match(/^\d+[GL]L$/) && val >= 20 && val <= 120) {
          purlinRequiredSpacing[key] = val;
        }
      }
    }
  }

  if (Object.keys(girtSpacingByWind).length === 0) {
    for (let r = 0; r < Math.min(25, data.length); r++) {
      const row = data[r];
      if (!row) continue;
      for (let c = 30; c < row.length - 1; c++) {
        const wind = num(row[c]);
        const spacing = num(row[c + 1]);
        if ([90, 110, 120, 130].includes(wind) && spacing >= 20 && spacing <= 80) {
          girtSpacingByWind[String(wind)] = spacing;
        }
      }
    }
  }

  if (Object.keys(originalGirtsByHeight).length === 0) {
    for (let r = 0; r < Math.min(25, data.length); r++) {
      const row = data[r];
      if (!row) continue;
      for (let c = 35; c < row.length - 1; c++) {
        const height = num(row[c]);
        const count = num(row[c + 1]);
        if (height >= 8 && height <= 20 && count >= 2 && count <= 10) {
          originalGirtsByHeight[String(height)] = count;
        }
      }
    }
  }

  return {
    purlinSpacing,
    trussCountByLength,
    trussSpacingByLength,
    verticalCountByWidth,
    verticalSpacingByWidth,
    verticalSpacingByWind,
    purlinRequiredSpacing,
    originalPurlinByWidth,
    girtSpacingByWind,
    originalGirtsByHeight,
  };
}

/**
 * Parse widespan "Snow Load Calculation" sheet.
 * This is a calculation sheet, not a lookup matrix.
 * We extract the pricing constants used:
 *   - Extra truss price by width group
 *   - Purlin L/FT cost
 *   - Vertical L/FT cost
 *   - Leg truss L/FT cost
 *   - Diagonal brace price ($350)
 *   - Girt perimeter cost
 */
export function readWidespanSnowCalc(ws: WorkSheet): {
  trussPriceByWidthGroup: PricingLookup;
  purlinCostPerFt: number;
  verticalCostPerFt: number;
  legTrussCostPerFt: number;
  diagonalBracePrice: number;
  girtCostPerFt: number;
} {
  const data = sheetToArray(ws);

  const trussPriceByWidthGroup: PricingLookup = {};
  let purlinCostPerFt = 6;
  let verticalCostPerFt = 18;
  let legTrussCostPerFt = 90;
  let diagonalBracePrice = 350;
  let girtCostPerFt = 6;

  // Scan for truss price table (typically cols Z-AO, rows 1-2)
  // Width groups: 32-40 = 1605, 42-48 = 1865, 50-60 = 2145
  for (let r = 0; r < Math.min(10, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 20; c < row.length - 1; c++) {
      const width = num(row[c]);
      const price = num(row[c + 1]);
      if (width >= 32 && width <= 60 && price >= 1000 && price <= 5000) {
        trussPriceByWidthGroup[String(width)] = price;
      }
    }
  }

  // Scan for cost constants
  for (let r = 0; r < Math.min(30, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const label = cleanHeader(row[c]);
      if (label.toLowerCase().includes("purlin") && label.toLowerCase().includes("l/ft")) {
        const val = num(row[c + 1]) || num(row[c - 1]);
        if (val > 0 && val < 50) purlinCostPerFt = val;
      }
      if (label.toLowerCase().includes("vertical") && label.toLowerCase().includes("l/ft")) {
        const val = num(row[c + 1]) || num(row[c - 1]);
        if (val > 0 && val < 50) verticalCostPerFt = val;
      }
      if (label.toLowerCase().includes("leg") && label.toLowerCase().includes("truss")) {
        const val = num(row[c + 1]) || num(row[c - 1]);
        if (val > 0 && val < 200) legTrussCostPerFt = val;
      }
      if (label.toLowerCase().includes("diagonal") || label.toLowerCase().includes("db")) {
        const val = num(row[c + 1]) || num(row[c - 1]);
        if (val >= 300 && val <= 500) diagonalBracePrice = val;
      }
    }
  }

  return {
    trussPriceByWidthGroup,
    purlinCostPerFt,
    verticalCostPerFt,
    legTrussCostPerFt,
    diagonalBracePrice,
    girtCostPerFt,
  };
}
