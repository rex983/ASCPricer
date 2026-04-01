import type {
  BuildingConfig,
  PriceBreakdown,
  WidespanMatrices,
} from "@/types/pricing";
import { resolveWidespanKeys } from "./changers";
import { lookupMatrix, lookupValue } from "./lookups";
import {
  SHEET_METAL_MULTIPLIERS,
  WIDESPAN_PLANS_LEG_SURCHARGE,
  WIDESPAN_PLANS_DOOR_OPENING_COST,
  WIDESPAN_PLANS_OVER14_SURCHARGE,
  WIDESPAN_ANCHOR_PRICES,
} from "./constants";
import {
  calculateWidespanSnowEngineering,
} from "./snow-engineering";

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
  // Sheet metal multiplier applied (Changers B30 × Base Price)
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
  // Spreadsheet: price × sheetMetal × qtyMultiplier (0→0, 1→0.5, 2→1.0)
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
    // Matrix price is for both sides; qty 1 = half, qty 2 = full
    const qtyMultiplier = config.sidesQty === 1 ? 0.5 : 1.0;
    sides = sidePrice * sheetMetalMult * qtyMultiplier;
  }

  // ── Ends ──
  // Sheet metal multiplier applied, multiplied by qty
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
  // Ends use base price (no header beam needed)
  let rollUpDoorsEnds = 0;
  if (config.rollUpEndSize1 && config.rollUpEndQty1 > 0) {
    rollUpDoorsEnds +=
      lookupValue(matrices.accessories.rollUpDoors, config.rollUpEndSize1) *
      config.rollUpEndQty1;
  }
  if (config.rollUpEndSize2 && config.rollUpEndQty2 > 0) {
    rollUpDoorsEnds +=
      lookupValue(matrices.accessories.rollUpDoors, config.rollUpEndSize2) *
      config.rollUpEndQty2;
  }

  // Sides use with-header price (header beam included in spreadsheet col D-E)
  // Fall back to base + rollUpSideHeader if with-header lookup is empty
  const sidesLookup = Object.keys(matrices.accessories.rollUpDoorsWithHeader).length > 0
    ? matrices.accessories.rollUpDoorsWithHeader
    : null;

  let rollUpDoorsSides = 0;
  if (config.rollUpSideSize1 && config.rollUpSideQty1 > 0) {
    const price = sidesLookup
      ? lookupValue(sidesLookup, config.rollUpSideSize1)
      : lookupValue(matrices.accessories.rollUpDoors, config.rollUpSideSize1) +
        matrices.accessories.rollUpSideHeader;
    rollUpDoorsSides += price * config.rollUpSideQty1;
  }
  if (config.rollUpSideSize2 && config.rollUpSideQty2 > 0) {
    const price = sidesLookup
      ? lookupValue(sidesLookup, config.rollUpSideSize2)
      : lookupValue(matrices.accessories.rollUpDoors, config.rollUpSideSize2) +
        matrices.accessories.rollUpSideHeader;
    rollUpDoorsSides += price * config.rollUpSideQty2;
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
  // Spreadsheet: sides × qtyMultiplier (0→0, 1→0.5, 2→1.0), ends × endsQty
  let wainscot = 0;
  if (config.wainscot === "full" || config.wainscot === "sides") {
    const sideWainscot = lookupValue(matrices.wainscot.sides, String(keys.length));
    const sidesQtyMult = config.sidesQty === 0 ? 0 : config.sidesQty === 1 ? 0.5 : 1.0;
    wainscot += sideWainscot * sidesQtyMult;
  }
  if (config.wainscot === "full" || config.wainscot === "ends") {
    const endWainscot = lookupValue(matrices.wainscot.ends, String(keys.width));
    wainscot += endWainscot * config.endsQty;
  }

  // ── Snow Engineering ──
  const engineeringResult = calculateWidespanSnowEngineering(
    config,
    matrices,
    keys
  );
  const snowEngineering = engineeringResult.cost;

  // No diagonal bracing for widespan buildings
  const diagonalBracing = 0;

  // ── Anchors ──
  // Spreadsheet: anchor count = (totalTrusses × 8) + (totalVerticals × 2)
  // Uses total counts (max of needed vs original), not extras.
  // Anchor types: titen_hd ($15/ea), concrete ($0/ea), or none
  let anchors = 0;
  if (config.anchorType && config.anchorType !== "none") {
    const anchorCostEach = WIDESPAN_ANCHOR_PRICES[config.anchorType] ?? 0;
    if (anchorCostEach > 0) {
      const anchorCount = Math.max(
        0,
        engineeringResult.totalTrusses * 8 + engineeringResult.totalVerticals * 2
      );
      anchors = anchorCount * anchorCostEach;
    }
  }

  // ── Plans & Calculations (separate from building price) ──
  let plans = 0;
  let calculations = 0;
  if (config.includePlans) {
    // Base plans cost from matrix
    const basePlans = lookupMatrix(matrices.plans, String(keys.width), String(keys.length));

    // Leg height surcharge (height 13+ adds $150-$575)
    const legSurcharge = WIDESPAN_PLANS_LEG_SURCHARGE[String(config.height)] ?? 0;

    // Door opening cost (3+ roll-up doors adds $100-$275)
    const totalDoorOpenings =
      config.rollUpEndQty1 + config.rollUpEndQty2 +
      config.rollUpSideQty1 + config.rollUpSideQty2;
    const doorOpeningCost = WIDESPAN_PLANS_DOOR_OPENING_COST[
      String(Math.min(totalDoorOpenings, 10))
    ] ?? 0;

    // Over-14' door surcharge ($100 per door over 14' tall)
    let over14Count = 0;
    const checkOver14 = (size?: string, qty?: number) => {
      if (!size || !qty) return;
      const parts = size.split("x").map(Number);
      if (parts.length === 2 && (parts[0] >= 16 || parts[1] >= 16)) {
        over14Count += qty;
      }
    };
    checkOver14(config.rollUpEndSize1, config.rollUpEndQty1);
    checkOver14(config.rollUpEndSize2, config.rollUpEndQty2);
    checkOver14(config.rollUpSideSize1, config.rollUpSideQty1);
    checkOver14(config.rollUpSideSize2, config.rollUpSideQty2);
    const over14Surcharge = over14Count * WIDESPAN_PLANS_OVER14_SURCHARGE;

    plans = basePlans + legSurcharge + doorOpeningCost + over14Surcharge;

    // Calculations cost from matrix
    calculations = lookupMatrix(matrices.calculations, String(keys.width), String(keys.length));
  }

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
    insulation + wainscot + snowEngineering + diagonalBracing + anchors;

  const taxAmount = Math.round(subtotal * config.taxRate * 100) / 100;
  const total = subtotal + taxAmount + laborEquipment;

  return {
    basePrice, roofStyle: 0, legs, sides, ends,
    walkInDoors, windows, rollUpDoorsEnds, rollUpDoorsSides,
    insulation, snowEngineering, snowEngineeringBreakdown: engineeringResult.breakdown, diagonalBracing,
    anchors, wainscot, plans, calculations,
    subtotal, laborEquipment,
    taxRate: config.taxRate, taxAmount, total,
  };
}
