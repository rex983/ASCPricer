import type {
  BuildingConfig,
  PriceBreakdown,
  WidespanMatrices,
} from "@/types/pricing";
import { resolveWidespanKeys } from "./changers";
import { lookupMatrix, lookupValue } from "./lookups";
import { SHEET_METAL_MULTIPLIERS } from "./constants";
import {
  calculateWidespanSnowEngineering,
  isDiagonalBracingNeeded,
} from "./snow-engineering";
import {
  WIDESPAN_BRACE_COUNT_SHORT,
  WIDESPAN_BRACE_COUNT_LONG,
  WIDESPAN_BRACE_ENDS_EXTRA,
} from "./constants";

/**
 * Calculate full price breakdown for a widespan building (width 32-60).
 */
export function calculateWidespanPrice(
  config: BuildingConfig,
  matrices: WidespanMatrices
): PriceBreakdown {
  const keys = resolveWidespanKeys(config, matrices);
  const sheetMetalMult =
    SHEET_METAL_MULTIPLIERS[config.sheetMetal || "29g_agg"];

  // ── Base Price ──
  // Widespan lengths 20-200 are all hardcoded (no extrapolation)
  const basePrice =
    lookupMatrix(matrices.basePrice, keys.basePriceKey, String(keys.length)) *
    sheetMetalMult;

  // ── Legs ──
  const legs = lookupMatrix(
    matrices.legs,
    String(config.height),
    String(keys.length)
  );

  // ── Sides ──
  // Widespan sides are looked up by panel count (not height like standard)
  // The panel count corresponds to height in 1ft increments
  let sides = 0;
  if (config.sidesQty > 0 && config.sidesCoverage !== "open") {
    let panelCount: string;
    if (config.sidesCoverage === "fully_enclosed") {
      panelCount = String(config.height);
    } else {
      const match = config.sidesCoverage.match(/(\d+)/);
      panelCount = match ? match[1] : String(config.height);
    }

    const sidePrice = lookupMatrix(
      matrices.sides,
      panelCount,
      String(keys.length)
    );
    sides = sidePrice * sheetMetalMult; // matrix gives total for all enclosed sides
  }

  // ── Ends ──
  let ends = 0;
  if (config.endsQty > 0) {
    let endTypeCode: string;
    if (config.endType === "enclosed") endTypeCode = "FE";
    else if (config.endType === "gable") endTypeCode = "GE";
    else endTypeCode = "FE";

    const endColKey = `${keys.width}-${endTypeCode}`;
    const endPrice = lookupMatrix(
      matrices.ends,
      endColKey,
      String(config.height)
    );
    ends = endPrice * config.endsQty * sheetMetalMult;
  }

  // ── Walk-In Doors ──
  const walkInDoors = config.walkInDoorType
    ? lookupValue(matrices.accessories.walkInDoors, config.walkInDoorType) *
      config.walkInDoorQty
    : 0;

  // ── Windows ──
  const windows = config.windowType
    ? lookupValue(matrices.accessories.windows, config.windowType) *
      config.windowQty
    : 0;

  // ── Roll-Up Doors ──
  let rollUpDoorsEnds = 0;
  if (config.rollUpEndSize && config.rollUpEndQty > 0) {
    rollUpDoorsEnds =
      lookupValue(matrices.accessories.rollUpDoors, config.rollUpEndSize) *
      config.rollUpEndQty;
  }

  let rollUpDoorsSides = 0;
  if (config.rollUpSideSize && config.rollUpSideQty > 0) {
    const baseRollUp = lookupValue(
      matrices.accessories.rollUpDoors,
      config.rollUpSideSize
    );
    rollUpDoorsSides =
      (baseRollUp + matrices.accessories.rollUpSideHeader) *
      config.rollUpSideQty;
  }

  // ── Insulation ──
  let insulation = 0;
  if (config.insulationType !== "none" && config.insulationScope !== "none") {
    const rate =
      config.insulationType === "fiberglass"
        ? matrices.insulation.fiberglassRate
        : matrices.insulation.thermalRate;

    const roofRaw = (keys.width + 3) * keys.length * rate;

    if (config.insulationScope === "roof_only") {
      insulation = Math.ceil(roofRaw / 10) * 10;
    } else {
      const sideRaw = (config.height + 2) * keys.length * config.sidesQty * rate;
      const endRaw = (config.height + 3) * keys.width * config.endsQty * rate;
      insulation = Math.round((roofRaw + sideRaw + endRaw) / 10) * 10;
    }
  }

  // ── Wainscot ──
  let wainscot = 0;
  if (config.wainscot === "full" || config.wainscot === "sides") {
    wainscot += lookupValue(matrices.wainscot.sides, String(keys.length));
  }
  if (config.wainscot === "full" || config.wainscot === "ends") {
    wainscot += lookupValue(matrices.wainscot.ends, String(keys.width));
  }

  // ── Snow Engineering ──
  const snowEngineering = calculateWidespanSnowEngineering(
    config,
    matrices,
    keys
  );

  // ── Diagonal Bracing (automatic — 3-trigger system) ──
  const dbNeeded = isDiagonalBracingNeeded(config, false);

  let diagonalBracing = 0;
  if (dbNeeded) {
    const bracePrice = matrices.snow.diagonalBracePrice || 350;
    let braceCount =
      keys.length <= 50 ? WIDESPAN_BRACE_COUNT_SHORT : WIDESPAN_BRACE_COUNT_LONG;

    // Add extra braces for enclosed ends
    if (config.endType === "gable" || config.endType === "enclosed") {
      braceCount += WIDESPAN_BRACE_ENDS_EXTRA * config.endsQty;
    }

    diagonalBracing = braceCount * bracePrice;
  }

  // ── Plans & Calculations (separate from building price) ──
  // Widespan "Plans & Calcs" sheet — plans is a simple lookup, no surcharges yet
  const plans = config.includePlans
    ? lookupMatrix(matrices.plans, String(keys.width), String(keys.length))
    : 0;
  const calculations = 0; // TODO: parse widespan calcs matrix when available

  // ── Labor/Equipment ──
  const laborEquipment = lookupMatrix(
    matrices.laborEquipment,
    String(keys.width),
    String(keys.length)
  );

  // ── Totals ──
  const subtotal =
    basePrice + legs + sides + ends +
    walkInDoors + windows + rollUpDoorsEnds + rollUpDoorsSides +
    insulation + wainscot + snowEngineering + diagonalBracing;

  const taxAmount = Math.round(subtotal * config.taxRate * 100) / 100;
  const total = subtotal + taxAmount + laborEquipment;

  return {
    basePrice, roofStyle: 0, legs, sides, ends,
    walkInDoors, windows, rollUpDoorsEnds, rollUpDoorsSides,
    insulation, snowEngineering, diagonalBracing,
    anchors: 0, wainscot, plans, calculations,
    subtotal, laborEquipment,
    taxRate: config.taxRate, taxAmount, total,
  };
}
