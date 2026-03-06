import type { WorkSheet } from "xlsx";
import { utils as XLSXUtils } from "xlsx";
import type { PricingMatrix, PricingLookup } from "@/types/pricing";

/** Read a sheet as a 2D array of raw values */
export function sheetToArray(ws: WorkSheet): (string | number)[][] {
  return XLSXUtils.sheet_to_json(ws, { header: 1, defval: "" }) as (
    | string
    | number
  )[][];
}

/** Parse a number from a cell value, returning 0 for non-numeric */
export function num(v: unknown): number {
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const n = parseFloat(v.replace(/[$,]/g, ""));
    return isNaN(n) ? 0 : n;
  }
  return 0;
}

/** Clean a header string: trim whitespace, normalize */
export function cleanHeader(v: unknown): string {
  return String(v ?? "")
    .trim()
    .replace(/\s+/g, " ");
}

/**
 * Read a standard 2D matrix from a sheet.
 * - Row 0 = column headers (col keys)
 * - Column A = row keys
 * - Data starts at (startRow, startCol)
 *
 * Returns matrix[colHeader][rowKey] for easy lookup by config key then length.
 * If transpose=true, returns matrix[rowKey][colHeader] instead.
 */
export function readMatrix(
  data: (string | number)[][],
  opts: {
    headerRow?: number; // default 0
    dataStartRow?: number; // default 1
    dataEndRow?: number; // reads until empty row key
    rowKeyCol?: number; // default 0
    dataStartCol?: number; // default 1
    dataEndCol?: number; // reads until empty col header
    transpose?: boolean; // default false: matrix[colKey][rowKey]
  } = {}
): PricingMatrix {
  const {
    headerRow = 0,
    dataStartRow = 1,
    rowKeyCol = 0,
    dataStartCol = 1,
    transpose = false,
  } = opts;

  const headers = data[headerRow] || [];
  const matrix: PricingMatrix = {};

  // Determine column range
  const endCol =
    opts.dataEndCol ??
    (() => {
      let last = dataStartCol;
      for (let c = dataStartCol; c < headers.length; c++) {
        const h = cleanHeader(headers[c]);
        if (h && h !== "0" && h !== "") last = c + 1;
      }
      return last;
    })();

  for (
    let r = dataStartRow;
    r < (opts.dataEndRow ?? data.length);
    r++
  ) {
    const row = data[r];
    if (!row) break;
    const rowKey = cleanHeader(row[rowKeyCol]);
    if (!rowKey || rowKey === "0") continue;

    for (let c = dataStartCol; c < endCol; c++) {
      const colKey = cleanHeader(headers[c]);
      if (!colKey || colKey === "0") continue;

      const value = num(row[c]);
      if (transpose) {
        if (!matrix[rowKey]) matrix[rowKey] = {};
        matrix[rowKey][colKey] = value;
      } else {
        if (!matrix[colKey]) matrix[colKey] = {};
        matrix[colKey][rowKey] = value;
      }
    }
  }

  return matrix;
}

/**
 * Read a simple key-value lookup from two columns.
 */
export function readLookup(
  data: (string | number)[][],
  opts: {
    startRow: number;
    endRow?: number;
    keyCol: number;
    valueCol: number;
  }
): PricingLookup {
  const lookup: PricingLookup = {};
  for (
    let r = opts.startRow;
    r < (opts.endRow ?? data.length);
    r++
  ) {
    const row = data[r];
    if (!row) break;
    const key = cleanHeader(row[opts.keyCol]);
    if (!key) break;
    lookup[key] = num(row[opts.valueCol]);
  }
  return lookup;
}

/**
 * Read a horizontal lookup from a single row: maps column headers to values.
 */
export function readRowLookup(
  data: (string | number)[][],
  opts: {
    headerRow: number;
    valueRow: number;
    startCol: number;
    endCol?: number;
  }
): PricingLookup {
  const lookup: PricingLookup = {};
  const headers = data[opts.headerRow] || [];
  const values = data[opts.valueRow] || [];
  const endCol = opts.endCol ?? headers.length;
  for (let c = opts.startCol; c < endCol; c++) {
    const key = cleanHeader(headers[c]);
    if (!key) continue;
    lookup[key] = num(values[c]);
  }
  return lookup;
}
