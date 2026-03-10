import type {
  BuildingConfig,
  PriceBreakdown,
  StandardMatrices,
} from "@/types/pricing";
import { resolveStandardKeys } from "./changers";
import { lookupWithLengthExtrapolation } from "./length-extrapolation";
import { lookupMatrix, lookupValue } from "./lookups";
import { getStandardBuildingType } from "./building-type";
import { calculateStandardSnowEngineering } from "./snow-engineering";
import {
  STANDARD_BRACE_BASE_PRICE,
  STANDARD_BRACE_TALL_SURCHARGE,
} from "./constants";

/**
 * Calculate full price breakdown for a standard building (width ≤ 30).
 */
export function calculateStandardPrice(
  config: BuildingConfig,
  matrices: StandardMatrices
): PriceBreakdown {
  const keys = resolveStandardKeys(config, matrices);
  const buildingType = getStandardBuildingType(keys.width, config.height);

  // ── Base Price ──
  const basePrice = lookupWithLengthExtrapolation(
    matrices.basePrice,
    keys.basePriceKey,
    keys.length
  );

  // ── Roof Style ──
  const roofStyle =
    config.roofStyle === "standard"
      ? 0
      : lookupWithLengthExtrapolation(
          matrices.roofStyle,
          keys.roofStyleKey,
          keys.length
        );

  // ── Legs ──
  const legMatrix =
    keys.width <= 24 ? matrices.legs.small : matrices.legs.large;
  const legs = lookupWithLengthExtrapolation(
    legMatrix,
    String(config.height),
    keys.length
  );

  // ── Sides ──
  let sides = 0;
  if (config.sidesQty > 0 && config.sidesCoverage !== "open") {
    let sideHeight: string;
    if (config.sidesCoverage === "fully_enclosed") {
      sideHeight = String(config.height);
    } else {
      const match = config.sidesCoverage.match(/(\d+)/);
      sideHeight = match ? match[1] : String(config.height);
    }

    const orientKey = keys.sidesOrientationKey;
    const lengthOrientKey = `${keys.length}-${orientKey}`;
    const sidePrice = lookupMatrix(matrices.sides, lengthOrientKey, sideHeight);
    sides = sidePrice; // price is total for all enclosed sides

    // V-side surcharge
    if (orientKey === "V" && Object.keys(matrices.vSidesSurcharge).length > 0) {
      for (const [label, lookup] of Object.entries(matrices.vSidesSurcharge)) {
        const h = config.height;
        if (
          (label.includes("6") && label.includes("10") && h >= 6 && h <= 10) ||
          (label.includes("11") && label.includes("15") && h >= 11 && h <= 15) ||
          (label.includes("16") && label.includes("20") && h >= 16 && h <= 20)
        ) {
          sides += (lookup[String(keys.length)] ?? 0) * config.sidesQty;
        }
      }
    }
  }

  // ── Ends ──
  let ends = 0;
  if (config.endsQty > 0) {
    const orientKey = keys.endsOrientationKey;
    let endTypeCode: string;
    if (config.endType === "enclosed") endTypeCode = "FE";
    else if (config.endType === "gable") endTypeCode = "G";
    else endTypeCode = "EG";

    const endColKey = `${keys.width}-${orientKey}-${endTypeCode}`;
    const endPrice = lookupMatrix(matrices.ends, endColKey, String(config.height));
    ends = endPrice * config.endsQty;

    // V-end surcharge
    if (orientKey === "V" && Object.keys(matrices.vEndsSurcharge).length > 0) {
      for (const [label, lookup] of Object.entries(matrices.vEndsSurcharge)) {
        const h = config.height;
        if (
          (label.includes("6") && label.includes("10") && h >= 6 && h <= 10) ||
          (label.includes("11") && label.includes("15") && h >= 11 && h <= 15) ||
          (label.includes("16") && label.includes("20") && h >= 16 && h <= 20)
        ) {
          ends += (lookup[String(keys.width)] ?? 0) * config.endsQty;
        }
      }
    }
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
  if (config.insulationType !== "none") {
    const rate =
      config.insulationType === "fiberglass"
        ? matrices.insulation.fiberglassRate
        : matrices.insulation.thermalRate;

    const roofSqFt = keys.width * keys.length;
    const sideSqFt = config.height * keys.length * config.sidesQty;
    const endSqFt = config.height * keys.width * config.endsQty;
    const totalSqFt = roofSqFt + sideSqFt + endSqFt;
    insulation = Math.round((totalSqFt * rate) / 10) * 10;
  }

  // ── Snow Engineering ──
  const snowRaw = calculateStandardSnowEngineering(config, matrices, keys);
  const contactEngineer = snowRaw === -1; // -1 = beyond standard engineering
  const snowEngineering = contactEngineer ? 0 : snowRaw;

  // ── Diagonal Bracing ──
  let diagonalBracing = 0;
  if (config.diagonalBracing) {
    const bracePrice = matrices.snow.diagonalBracePrice || STANDARD_BRACE_BASE_PRICE;
    const tallSurcharge =
      config.height > 12
        ? matrices.snow.diagonalBraceTallSurcharge || STANDARD_BRACE_TALL_SURCHARGE
        : 0;

    let needsBracing = true;
    if (config.state && Object.keys(matrices.snow.windThresholdByState).length > 0) {
      const threshold = matrices.snow.windThresholdByState[config.state];
      if (threshold && config.windRating < threshold) {
        needsBracing = false;
      }
    }

    if (needsBracing) {
      const braceCount = keys.length <= 50 ? 8 : 10;
      diagonalBracing = braceCount * (bracePrice + tallSurcharge);
    }
  }

  // ── Plans ──
  const plans = config.includePlans
    ? lookupMatrix(matrices.plans, String(keys.width), String(keys.length))
    : 0;

  // ── Labor/Equipment ──
  const typeCode = buildingType === "M" ? "S" : buildingType;
  const laborKey = `${keys.width}${typeCode}`;
  const laborEquipment = lookupMatrix(matrices.laborEquipment, laborKey, String(keys.length));

  // ── Totals ──
  const subtotal =
    basePrice + roofStyle + legs + sides + ends +
    walkInDoors + windows + rollUpDoorsEnds + rollUpDoorsSides +
    insulation + snowEngineering + diagonalBracing + plans;

  const taxAmount = Math.round(subtotal * config.taxRate * 100) / 100;
  const total = subtotal + taxAmount + laborEquipment;

  return {
    basePrice, roofStyle, legs, sides, ends,
    walkInDoors, windows, rollUpDoorsEnds, rollUpDoorsSides,
    insulation, snowEngineering, contactEngineer, diagonalBracing,
    anchors: 0, wainscot: 0, plans,
    subtotal, laborEquipment,
    taxRate: config.taxRate, taxAmount, total,
  };
}
