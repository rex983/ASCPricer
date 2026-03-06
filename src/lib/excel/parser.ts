import * as XLSX from "xlsx";
import type { PricingMatrices } from "@/types/pricing";
import { detectSpreadsheetType, type SpreadsheetDetection } from "./detect-type";
import { parseStandardWorkbook } from "./parser-standard";
import { parseWidespanWorkbook } from "./parser-widespan";
import { validateMatrices, type ValidationResult } from "./validators";

export interface ParseResult {
  detection: SpreadsheetDetection;
  matrices: PricingMatrices;
  validation: ValidationResult;
}

/**
 * Main entry point: parse an Excel spreadsheet buffer into pricing matrices.
 *
 * 1. Read the workbook
 * 2. Auto-detect standard vs widespan
 * 3. Parse all sheets into structured matrices
 * 4. Validate the result
 */
export function parseSpreadsheet(buffer: ArrayBuffer | Uint8Array): ParseResult {
  const workbook = XLSX.read(buffer, { type: "array" });
  const detection = detectSpreadsheetType(workbook);

  let matrices: PricingMatrices;
  if (detection.type === "standard") {
    matrices = parseStandardWorkbook(workbook);
  } else {
    matrices = parseWidespanWorkbook(workbook);
  }

  const validation = validateMatrices(matrices);

  return { detection, matrices, validation };
}
