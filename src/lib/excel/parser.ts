import * as XLSX from "xlsx";
import type { PricingMatrices } from "@/types/pricing";
import { detectSpreadsheetType, type SpreadsheetDetection } from "./detect-type";
import { parseStandardWorkbook } from "./parser-standard";
import { parseWidespanWorkbook } from "./parser-widespan";
import { validateMatrices, type ValidationResult } from "./validators";
import { FIBERGLASS_RATE, THERMAL_RATE } from "@/lib/pricing/constants";

/** Insulation rates ($/sqft) used where the spreadsheet doesn't supply them. */
export interface InsulationDefaults {
  fiberglass: number;
  thermal: number;
}

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
export function parseSpreadsheet(
  buffer: ArrayBuffer | Uint8Array,
  filename?: string,
  insulation: InsulationDefaults = { fiberglass: FIBERGLASS_RATE, thermal: THERMAL_RATE }
): ParseResult {
  const workbook = XLSX.read(buffer, { type: "array" });
  const detection = detectSpreadsheetType(workbook, filename);

  let matrices: PricingMatrices;
  if (detection.type === "standard") {
    matrices = parseStandardWorkbook(workbook, insulation);
  } else {
    matrices = parseWidespanWorkbook(workbook, insulation);
  }

  const validation = validateMatrices(matrices);

  return { detection, matrices, validation };
}
