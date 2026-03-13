import type { WorkBook } from "xlsx";
import { utils as XLSXUtils } from "xlsx";
import type { StandardMatrices } from "@/types/pricing";
import { readBasePrice } from "./sheet-readers/standard-base";
import { readRoofStyle } from "./sheet-readers/standard-roof-style";
import { readLegs } from "./sheet-readers/standard-legs";
import { readSides } from "./sheet-readers/standard-sides";
import { readEnds } from "./sheet-readers/standard-ends";
import {
  readAccessories,
  readStandardRollUpDoors,
} from "./sheet-readers/standard-accessories";
import { readAnchors } from "./sheet-readers/standard-anchors";
import { readLaborEquipment } from "./sheet-readers/standard-labor";
import {
  readPlans,
  readCalculations,
  readPlansLegSurcharge,
  readPlansDoorOpeningCost,
} from "./sheet-readers/standard-plans";
import { readChangers } from "./sheet-readers/standard-changers";
import {
  readTrussSpacing,
  readTrussCounts,
  readHatChannels,
  readGirtSpacing as readGirtsSheet,
  readVerticals,
  readDiagonalBracing,
  readSnowChangers,
} from "./sheet-readers/standard-snow";

/** Find a sheet by name, tolerating trailing spaces in sheet names */
function findSheet(workbook: WorkBook, name: string) {
  // Try exact match first
  if (workbook.Sheets[name]) return workbook.Sheets[name];
  // Try with/without trailing space
  if (workbook.Sheets[name + " "]) return workbook.Sheets[name + " "];
  const trimmed = name.trim();
  if (workbook.Sheets[trimmed]) return workbook.Sheets[trimmed];
  // Fuzzy: find first sheet whose trimmed name matches
  for (const key of Object.keys(workbook.Sheets)) {
    if (key.trim() === trimmed) return workbook.Sheets[key];
  }
  return null;
}

function getSheet(workbook: WorkBook, name: string) {
  const ws = findSheet(workbook, name);
  if (!ws) throw new Error(`Sheet "${name}" not found in workbook`);
  return ws;
}

function tryGetSheet(workbook: WorkBook, name: string) {
  return findSheet(workbook, name);
}

/**
 * Parse a standard spreadsheet (21 sheets, width ≤ 30) into pricing matrices.
 */
export function parseStandardWorkbook(workbook: WorkBook): StandardMatrices {
  // Core pricing sheets
  const basePrice = readBasePrice(getSheet(workbook, "Pricing - Base"));
  const roofStyle = readRoofStyle(getSheet(workbook, "Pricing - Roof Style"));
  const { small, large } = readLegs(getSheet(workbook, "Pricing - Legs"));
  const { sides, vSidesSurcharge } = readSides(
    getSheet(workbook, "Pricing - Sides")
  );
  const { ends, vEndsSurcharge } = readEnds(
    getSheet(workbook, "Pricing - Ends")
  );

  // Accessories
  const { walkInDoors, windows } = readAccessories(
    getSheet(workbook, "Pricing - Accessories")
  );
  const rollUpDoors = readStandardRollUpDoors(
    getSheet(workbook, "Pricing - Accessories")
  );

  // Anchors, labor, plans
  const anchors = readAnchors(getSheet(workbook, "Pricing - Anchors"));
  const laborEquipment = readLaborEquipment(
    getSheet(workbook, "Pricing - Labor-EQ")
  );
  const plansSheet = getSheet(workbook, "Plans for Buildings");
  const plans = readPlans(plansSheet);
  const calculations = readCalculations(plansSheet);
  const plansLegSurcharge = readPlansLegSurcharge(plansSheet);
  const plansDoorOpeningCost = readPlansDoorOpeningCost(plansSheet);

  // Changers
  const changers = readChangers(getSheet(workbook, "Pricing - Changers"));

  // Snow/engineering
  const trussSpacingSheet = tryGetSheet(workbook, "Snow - Truss Spacing");
  const trussCountsSheet = tryGetSheet(workbook, "Snow - Trusses");
  const hatChannelsSheet = tryGetSheet(workbook, "Snow - Hat Channels");
  const girtsSheet = tryGetSheet(workbook, "Snow - Girts");
  const verticalsSheet = tryGetSheet(workbook, "Snow - Verticals");
  const dbSheet = tryGetSheet(workbook, "Snow - Diagonal Bracing");
  const snowChangersSheet = tryGetSheet(workbook, "Snow - Changers");

  const trussSpacing = trussSpacingSheet
    ? readTrussSpacing(trussSpacingSheet)
    : {};
  const trussCounts = trussCountsSheet
    ? readTrussCounts(trussCountsSheet)
    : {};
  const { spacing: hatChannelSpacing, originalCounts: hatChannelCounts } =
    hatChannelsSheet ? readHatChannels(hatChannelsSheet) : { spacing: {}, originalCounts: {} };
  const { spacing: girtSpacing, girtCountsByHeight } = girtsSheet
    ? readGirtsSheet(girtsSheet)
    : { spacing: {}, girtCountsByHeight: {} };
  const {
    spacing: verticalSpacing,
    originalCounts: verticalCounts,
  } = verticalsSheet
    ? readVerticals(verticalsSheet)
    : { spacing: {}, originalCounts: {} };

  const {
    windThresholdByState,
    baseBracePrice: diagonalBracePrice,
    tallSurcharge: diagonalBraceTallSurcharge,
  } = dbSheet
    ? readDiagonalBracing(dbSheet)
    : { windThresholdByState: {}, baseBracePrice: 90, tallSurcharge: 50 };

  const snowChangers = snowChangersSheet
    ? readSnowChangers(snowChangersSheet)
    : {
        windLoadBuckets: {},
        snowLoadOptions: {},
        heightClassification: {},
        feetUsedByHeight: {},
        pieTrussPrice: {},
        trussPriceByWidthState: {},
        channelPriceByState: {},
        tubingPriceByState: {},
      };

  // Parse permit flag from Quote Sheet B27 (region-level, "Yes" or "No")
  const quoteSheet = tryGetSheet(workbook, "Quote Sheet");
  let permitRequired = false;
  if (quoteSheet) {
    const cellRef = quoteSheet["B27"];
    const val = cellRef ? String(cellRef.v ?? "").trim().toLowerCase() : "";
    permitRequired = val === "yes";
  }

  return {
    type: "standard",
    basePrice,
    roofStyle,
    legs: { small, large },
    sides,
    vSidesSurcharge,
    ends,
    vEndsSurcharge,
    accessories: {
      walkInDoors,
      windows,
      rollUpDoors,
      rollUpSideHeader: 260,
      rollUpLargeSize: 490,
    },
    insulation: {
      fiberglassRate: 2.25,
      thermalRate: 1.65,
    },
    anchors,
    laborEquipment,
    plans,
    calculations,
    plansSnowSurcharge: snowChangers.snowLoadOptions,
    plansLegSurcharge,
    plansDoorOpeningCost,
    snow: {
      trussSpacing,
      trussCounts,
      hatChannelSpacing,
      hatChannelCounts,
      girtSpacing,
      girtCountsByHeight,
      verticalSpacing,
      verticalCounts,
      trussPriceByWidthState: snowChangers.trussPriceByWidthState,
      channelPriceByState: snowChangers.channelPriceByState,
      tubingPriceByState: snowChangers.tubingPriceByState,
      windLoadBuckets: snowChangers.windLoadBuckets,
      heightClassification: snowChangers.heightClassification,
      feetUsedByHeight: snowChangers.feetUsedByHeight,
      pieTrussPrice: snowChangers.pieTrussPrice,
      diagonalBracePrice,
      diagonalBraceTallSurcharge,
      windThresholdByState,
      permitRequired,
    },
    changers,
  };
}
