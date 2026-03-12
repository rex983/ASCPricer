// ---- Building Configuration (user inputs) ----

export type SpreadsheetType = "standard" | "widespan";
export type RoofStyle = "standard" | "a_frame_horizontal" | "a_frame_vertical";
export type SidesCoverage = "open" | "fully_enclosed" | string; // string for "3'-20' Sides Down"
export type EndType = "enclosed" | "gable" | "extended_gable";
export type PanelOrientation = "horizontal" | "vertical";
export type Gauge = 12 | 14;
export type SheetMetal = "29g_agg" | "26g_agg" | "26g_pbr"; // widespan only
export type InsulationType = "fiberglass" | "thermal" | "none";

export interface BuildingConfig {
  // Core dimensions
  width: number; // 12-30 (standard) or 32-60 (widespan)
  length: number; // 20-100 (standard) or 20-200 (widespan)
  height: number; // 6-20 (standard) or 8-20 (widespan)
  gauge: Gauge;

  // Roof
  roofStyle: RoofStyle; // standard only; widespan always AFV

  // Sheet metal (widespan only)
  sheetMetal?: SheetMetal;

  // Sides
  sidesCoverage: SidesCoverage;
  sidesOrientation: PanelOrientation; // standard only
  sidesQty: 0 | 1 | 2;

  // Ends
  endType: EndType;
  endsOrientation: PanelOrientation; // standard only
  endsQty: 0 | 1 | 2;

  // Doors & Windows
  walkInDoorType?: string;
  walkInDoorQty: number;
  windowType?: string;
  windowQty: number;
  rollUpEndSize?: string;
  rollUpEndQty: number;
  rollUpSideSize?: string;
  rollUpSideQty: number;

  // Insulation
  insulationType: InsulationType;
  insulationQty: number; // number of sides/components to insulate

  // Wainscot (widespan only)
  wainscot?: "none" | "full" | "sides" | "ends";

  // Engineering
  snowLoad?: string; // e.g., "20 Roof Load", "30GL"
  windRating: number; // MPH
  diagonalBracing: boolean;

  // Plans & Anchors
  includePlans?: boolean;
  anchorType?: string;

  // State (for snow/wind engineering thresholds)
  state?: string;

  // Tax
  taxRate: number; // e.g., 0.0825 for 8.25%
}

// ---- Price Breakdown (output) ----

export interface PriceBreakdown {
  basePrice: number;
  roofStyle: number;
  legs: number;
  sides: number;
  ends: number;
  walkInDoors: number;
  windows: number;
  rollUpDoorsEnds: number;
  rollUpDoorsSides: number;
  insulation: number;
  snowEngineering: number;
  contactEngineer?: boolean; // true when snow load exceeds standard engineering
  diagonalBracing: number;
  anchors: number;
  wainscot: number; // widespan only
  plans: number;
  subtotal: number;
  laborEquipment: number;
  taxRate: number;
  taxAmount: number;
  total: number;
}

// ---- Pricing Matrices (parsed from spreadsheets) ----

/** A 2D lookup table: matrix[rowKey][colKey] = number */
export type PricingMatrix = Record<string, Record<string, number>>;

/** A simple lookup table: map[key] = number */
export type PricingLookup = Record<string, number>;

export interface StandardMatrices {
  type: "standard";
  basePrice: PricingMatrix; // [width-gauge][length] → price
  roofStyle: PricingMatrix; // [style-width][length] → price
  legs: {
    small: PricingMatrix; // width ≤ 24: [height][length] → price
    large: PricingMatrix; // width 26-30: [height][length] → price
  };
  sides: PricingMatrix; // [sideHeight][length-orientation] → price
  vSidesSurcharge: PricingMatrix; // vertical sides surcharge
  ends: PricingMatrix; // [width-endType][height] → price
  vEndsSurcharge: PricingMatrix; // vertical ends surcharge
  accessories: {
    walkInDoors: PricingLookup;
    windows: PricingLookup;
    rollUpDoors: PricingLookup;
    rollUpSideHeader: number; // $260
    rollUpLargeSize: number; // $490 for 10'+
  };
  insulation: {
    fiberglassRate: number; // 2.25
    thermalRate: number; // 1.65
  };
  anchors: PricingLookup; // [widthXendCount] → anchor count
  laborEquipment: PricingMatrix; // [width-type][length] → price
  plans: PricingMatrix; // [width][length] → price
  plansSnowSurcharge: PricingLookup; // snow load → surcharge
  snow: {
    trussSpacing: PricingMatrix; // [snowCode][configKey] → spacing
    trussCounts: PricingMatrix; // ["{width}-{state}"]["{length}"] → count
    hatChannelSpacing: PricingMatrix; // ["{bucketedTrussSpacing}-{snowCode}"][windSpeed] → spacing
    hatChannelCounts: PricingMatrix; // [stateWidth][count key] → count
    girtSpacing: PricingMatrix; // ["{bucketedTrussSpacing}"]["{windSpeed}"] → spacing
    girtCountsByHeight: PricingLookup; // height → original girt count
    verticalSpacing: PricingMatrix; // [height][windSpeed] → spacing
    verticalCounts: PricingLookup; // width → original count
    trussPriceByWidthState: PricingMatrix; // [state][width] → price
    channelPriceByState: PricingLookup; // state → $/ft ($2 or $2.50)
    tubingPriceByState: PricingLookup; // state → $/ft ($3-$4)
    windLoadBuckets: PricingLookup; // input MPH → category (105/115/130/140/155/165/180)
    heightClassification: PricingLookup; // height → S(0)/M(1)/T(2)
    feetUsedByHeight: PricingLookup; // height → feet used for truss leg surcharge
    pieTrussPrice: PricingLookup; // state → $/ft for truss leg surcharge ($15)
    diagonalBracePrice: number;
    diagonalBraceTallSurcharge: number;
    windThresholdByState: PricingLookup; // state → MPH threshold
    permitRequired: boolean; // region-level permit flag (parsed from Quote Sheet B27)
  };
  changers: {
    widthBuckets: PricingLookup; // raw width → bucket
    lengthBuckets: PricingLookup; // raw length → bucket
    heightToSidesKey: PricingLookup; // height → sides lookup key
    heightToLegsKey: PricingLookup; // height → legs lookup key
    gaugeLookup: PricingLookup;
    sheetMetalMultiplier: PricingLookup;
    buildingTypeByHeight: PricingLookup; // height → S/M/T/ET code
  };
}

export interface WidespanMatrices {
  type: "widespan";
  basePrice: PricingMatrix; // [width-gauge][length] → price
  legs: PricingMatrix; // [height][length] → price (single matrix)
  sides: PricingMatrix; // [sideHeight][length] → price
  ends: PricingMatrix; // [endType-width][height] → price
  accessories: {
    walkInDoors: PricingLookup;
    windows: PricingLookup;
    rollUpDoors: PricingLookup;
    rollUpSideHeader: number;
    rollUpLargeSize: number;
    rollUpOver14Surcharge: number;
  };
  insulation: {
    fiberglassRate: number;
    thermalRate: number;
  };
  wainscot: {
    sides: PricingLookup; // length → price
    ends: PricingLookup; // width → price
  };
  anchors: PricingLookup; // anchorType → per-anchor cost
  laborEquipment: PricingMatrix; // [widthGroup][lengthGroup] → price
  plans: PricingMatrix; // [widthGroup][length] → price
  snow: {
    trussSpacing: PricingMatrix;
    trussCounts: PricingMatrix;
    purlinSpacing: PricingMatrix;
    purlinCounts: PricingMatrix;
    girtSpacing: PricingMatrix;
    girtCounts: PricingMatrix;
    trussPriceByState: PricingLookup;
    diagonalBracePrice: number;
    verticalCountByWidth: PricingLookup; // width → count
    verticalSpacingByWind: PricingLookup; // wind → spacing inches
    purlinCostPerFt: number; // $6
    verticalCostPerFt: number; // $18
    girtCostPerFt: number; // $6
    legTrussCostPerFt: number; // $90
    windLoadMapping2: PricingLookup; // second wind mapping for girts: 90/110/120/130
  };
  changers: {
    widthBuckets: PricingLookup;
    lengthBuckets: PricingLookup;
    gaugeLookup: PricingLookup;
    sheetMetalMultiplier: PricingLookup;
    buildingTypeByWidthHeight: PricingMatrix; // [width][height] → S/G
    windLoadMapping: PricingLookup;
    snowLoadMapping: PricingLookup;
    heightClassification: PricingLookup;
  };
}

export type PricingMatrices = StandardMatrices | WidespanMatrices;
