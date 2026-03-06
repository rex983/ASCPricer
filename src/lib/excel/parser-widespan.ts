import type { WorkBook } from "xlsx";
import type { WidespanMatrices } from "@/types/pricing";
import { readWidespanBasePrice } from "./sheet-readers/widespan-base";
import { readWidespanLegs } from "./sheet-readers/widespan-legs";
import { readWidespanSides } from "./sheet-readers/widespan-sides";
import { readWidespanEnds } from "./sheet-readers/widespan-ends";
import {
  readWidespanDoorsWindows,
  readWidespanRollUpDoors,
} from "./sheet-readers/widespan-accessories";
import { readWidespanInsulationWainscot } from "./sheet-readers/widespan-insulation";
import { readWidespanEquipment } from "./sheet-readers/widespan-equipment";
import { readWidespanPlans } from "./sheet-readers/widespan-plans";
import { readWidespanChangers } from "./sheet-readers/widespan-changers";
import {
  readWidespanSnowLoad,
  readWidespanSnowCalc,
} from "./sheet-readers/widespan-snow";

function getSheet(workbook: WorkBook, name: string) {
  const ws = workbook.Sheets[name];
  if (!ws) throw new Error(`Sheet "${name}" not found in workbook`);
  return ws;
}

function tryGetSheet(workbook: WorkBook, name: string) {
  return workbook.Sheets[name] || null;
}

/**
 * Parse a widespan spreadsheet (14 sheets, width 32-60) into pricing matrices.
 */
export function parseWidespanWorkbook(workbook: WorkBook): WidespanMatrices {
  // Core pricing sheets
  const basePrice = readWidespanBasePrice(getSheet(workbook, "Base Price"));
  const legs = readWidespanLegs(getSheet(workbook, "Leg Height"));
  const sides = readWidespanSides(getSheet(workbook, "Sides"));
  const ends = readWidespanEnds(getSheet(workbook, "Ends"));

  // Accessories
  const { walkInDoors, windows } = readWidespanDoorsWindows(
    getSheet(workbook, "Doors - Windows")
  );
  const {
    rollUpDoors,
    headerSmall,
    headerLarge,
  } = readWidespanRollUpDoors(getSheet(workbook, "Roll Up Door"));

  // Insulation & wainscot
  const { fiberglassRate, thermalRate, wainscotSides, wainscotEnds } =
    readWidespanInsulationWainscot(
      getSheet(workbook, "Insulation - Wainscot")
    );

  // Equipment & plans
  const laborEquipment = readWidespanEquipment(
    getSheet(workbook, "Equipment")
  );
  const plans = readWidespanPlans(getSheet(workbook, "Plans & Calcs"));

  // Changers
  const changers = readWidespanChangers(getSheet(workbook, "Changers"));

  // Snow/engineering
  const snowLoadSheet = tryGetSheet(workbook, "Snow Load");
  const snowCalcSheet = tryGetSheet(workbook, "Snow Load Calculation");

  const snowLoad = snowLoadSheet
    ? readWidespanSnowLoad(snowLoadSheet)
    : {
        purlinSpacing: {},
        trussCountByLength: {},
        trussSpacingByLength: {},
        verticalCountByWidth: {},
        verticalSpacingByWidth: {},
        purlinRequiredSpacing: {},
        originalPurlinByWidth: {},
        girtSpacingByWind: {},
        originalGirtsByHeight: {},
      };

  const snowCalc = snowCalcSheet
    ? readWidespanSnowCalc(snowCalcSheet)
    : {
        trussPriceByWidthGroup: {},
        purlinCostPerFt: 6,
        verticalCostPerFt: 18,
        legTrussCostPerFt: 90,
        diagonalBracePrice: 350,
        girtCostPerFt: 6,
      };

  // Build structured snow matrices
  const trussSpacing: Record<string, Record<string, number>> = {};
  const trussCounts: Record<string, Record<string, number>> = {};
  for (const [len, count] of Object.entries(snowLoad.trussCountByLength)) {
    trussCounts[len] = { count };
  }
  for (const [len, spacing] of Object.entries(snowLoad.trussSpacingByLength)) {
    trussSpacing[len] = { spacing };
  }

  const purlinCounts: Record<string, Record<string, number>> = {};
  for (const [w, count] of Object.entries(snowLoad.originalPurlinByWidth)) {
    purlinCounts[w] = { count };
  }

  const girtSpacing: Record<string, Record<string, number>> = {};
  for (const [wind, spacing] of Object.entries(snowLoad.girtSpacingByWind)) {
    girtSpacing[wind] = { spacing };
  }
  const girtCounts: Record<string, Record<string, number>> = {};
  for (const [h, count] of Object.entries(snowLoad.originalGirtsByHeight)) {
    girtCounts[h] = { count };
  }

  return {
    type: "widespan",
    basePrice,
    legs,
    sides,
    ends,
    accessories: {
      walkInDoors,
      windows,
      rollUpDoors,
      rollUpSideHeader: headerSmall,
      rollUpLargeSize: headerLarge,
      rollUpOver14Surcharge: 0,
    },
    insulation: {
      fiberglassRate,
      thermalRate,
    },
    wainscot: {
      sides: wainscotSides,
      ends: wainscotEnds,
    },
    anchors: {},
    laborEquipment,
    plans,
    snow: {
      trussSpacing,
      trussCounts,
      purlinSpacing: snowLoad.purlinSpacing,
      purlinCounts,
      girtSpacing,
      girtCounts,
      trussPriceByState: snowCalc.trussPriceByWidthGroup,
      diagonalBracePrice: snowCalc.diagonalBracePrice,
    },
    changers,
  };
}
