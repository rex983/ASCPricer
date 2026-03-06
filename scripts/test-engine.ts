import * as fs from "fs";
import { parseSpreadsheet } from "../src/lib/excel/parser";
import { calculatePrice } from "../src/lib/pricing/engine";
import type { BuildingConfig } from "../src/types/pricing";

// ============================================
// TEST 1: Standard building from Quote Sheet
// 24w x 50L x 12H x 14G, AFV, Enclosed sides×2 HZ, Enclosed ends×2 HZ
// Expected: Base=4790, Roof=2325, Legs=685, Sides=1550, Ends=2800
//           Sales Total=12150, Tax 8.25%=1002.38, Total=13152.38
// ============================================

const stdBuffer = fs.readFileSync("C:/Users/Redir/spreadsheet.xlsx");
const stdResult = parseSpreadsheet(stdBuffer);

if (stdResult.matrices.type !== "standard") {
  throw new Error("Expected standard matrices");
}

const stdConfig: BuildingConfig = {
  width: 24,
  length: 51, // roof length=51, frame length rounds to 50
  height: 12,
  gauge: 14,
  roofStyle: "a_frame_vertical",
  sidesCoverage: "fully_enclosed",
  sidesOrientation: "horizontal",
  sidesQty: 2,
  endType: "enclosed",
  endsOrientation: "horizontal",
  endsQty: 2,
  walkInDoorType: undefined,
  walkInDoorQty: 0,
  windowType: undefined,
  windowQty: 0,
  rollUpEndSize: undefined,
  rollUpEndQty: 0,
  rollUpSideSize: undefined,
  rollUpSideQty: 0,
  insulationType: "none",
  insulationQty: 0,
  windRating: 90,
  diagonalBracing: false,
  taxRate: 0.0825,
};

const stdPrice = calculatePrice(stdConfig, stdResult.matrices);

console.log("=== STANDARD ENGINE TEST ===");
console.log("Config: 24w x 50L x 12H x 14G, AFV, Enclosed×2, Ends×2");
console.log("");
console.log("Component        | Calculated | Expected | Match");
console.log("─────────────────┼────────────┼──────────┼──────");
const checks = [
  ["Base Price", stdPrice.basePrice, 4790],
  ["Roof Style", stdPrice.roofStyle, 2325],
  ["Legs", stdPrice.legs, 685],
  ["Sides", stdPrice.sides, 1550],
  ["Ends", stdPrice.ends, 2800],
  ["Walk-In Doors", stdPrice.walkInDoors, 0],
  ["Windows", stdPrice.windows, 0],
  ["Roll-Up Ends", stdPrice.rollUpDoorsEnds, 0],
  ["Roll-Up Sides", stdPrice.rollUpDoorsSides, 0],
  ["Insulation", stdPrice.insulation, 0],
  ["Snow Eng", stdPrice.snowEngineering, 0],
  ["Diag Brace", stdPrice.diagonalBracing, 0],
  ["Plans", stdPrice.plans, 0],
  ["Labor/Equip", stdPrice.laborEquipment, 0],
  ["Subtotal", stdPrice.subtotal, 12150],
  ["Tax (8.25%)", stdPrice.taxAmount, 1002.38],
  ["Total", stdPrice.total, 13152.38],
];

let allPass = true;
for (const [label, calc, exp] of checks) {
  const match = Math.abs(Number(calc) - Number(exp)) < 1;
  if (!match) allPass = false;
  console.log(
    `${String(label).padEnd(17)}| ${String(calc).padStart(10)} | ${String(exp).padStart(8)} | ${match ? "✓" : "✗ MISMATCH"}`
  );
}

// ============================================
// TEST 2: Widespan building from Quote Sheet
// 60w x 50L x 10H x 12G, 29G Agg, Enclosed sides×2, Enclosed ends×1
// + door (48"x84" HD) ×1, windows (30"x53" DP) ×2, roll-up 8x10 ×1
// + Full Wainscot
// Expected: Base=27410, Sides=2770, Ends=5454, Door=815, Windows=500
//           Roll-up=985, Wainscot=1965
//           Sales Total=39899, Tax 6%=2393.94, Equipment=3850, Total=46142.94
// ============================================

console.log("\n=== WIDESPAN ENGINE TEST ===");

const wsBuffer = fs.readFileSync("C:/Users/Redir/spreadsheet-widespan.xlsx");
const wsResult = parseSpreadsheet(wsBuffer);

if (wsResult.matrices.type !== "widespan") {
  throw new Error("Expected widespan matrices");
}

const wsConfig: BuildingConfig = {
  width: 60,
  length: 51, // rounds to 50
  height: 10,
  gauge: 12,
  roofStyle: "a_frame_vertical",
  sheetMetal: "29g_agg",
  sidesCoverage: "fully_enclosed",
  sidesOrientation: "horizontal",
  sidesQty: 2,
  endType: "enclosed",
  endsOrientation: "horizontal",
  endsQty: 1,
  walkInDoorType: '48"x84" Heavy Duty Door',
  walkInDoorQty: 1,
  windowType: '30"x53" Double Pane Window',
  windowQty: 2,
  rollUpEndSize: "8x10",
  rollUpEndQty: 1,
  rollUpSideSize: undefined,
  rollUpSideQty: 0,
  insulationType: "none",
  insulationQty: 0,
  wainscot: "full",
  windRating: 90,
  diagonalBracing: false,
  taxRate: 0.06,
};

const wsPrice = calculatePrice(wsConfig, wsResult.matrices);

console.log("Config: 60w x 50L x 10H x 12G, 29G Agg, Sides×2, Ends×1 FE");
console.log("        + 48x84 HD Door, 30x53 DP Window×2, 8x10 Roll-up, Full Wainscot");
console.log("");
console.log("Component        | Calculated | Expected | Match");
console.log("─────────────────┼────────────┼──────────┼──────");

const wsChecks = [
  ["Base Price", wsPrice.basePrice, 27410],
  ["Legs", wsPrice.legs, 0],
  ["Sides", wsPrice.sides, 2770],
  ["Ends", wsPrice.ends, 5454],
  ["Walk-In Doors", wsPrice.walkInDoors, 815],
  ["Windows", wsPrice.windows, 500],
  ["Roll-Up Ends", wsPrice.rollUpDoorsEnds, 985],
  ["Roll-Up Sides", wsPrice.rollUpDoorsSides, 0],
  ["Insulation", wsPrice.insulation, 0],
  ["Wainscot", wsPrice.wainscot, 1965],
  ["Snow Eng", wsPrice.snowEngineering, 0],
  ["Diag Brace", wsPrice.diagonalBracing, 0],
  ["Plans", wsPrice.plans, 0],
  ["Subtotal", wsPrice.subtotal, 39899],
  ["Tax (6%)", wsPrice.taxAmount, 2393.94],
  ["Equipment", wsPrice.laborEquipment, 3850],
  ["Total", wsPrice.total, 46142.94],
];

for (const [label, calc, exp] of wsChecks) {
  const match = Math.abs(Number(calc) - Number(exp)) < 1;
  if (!match) allPass = false;
  console.log(
    `${String(label).padEnd(17)}| ${String(calc).padStart(10)} | ${String(exp).padStart(8)} | ${match ? "✓" : "✗ MISMATCH"}`
  );
}

console.log("\n" + (allPass ? "ALL TESTS PASSED ✓" : "SOME TESTS FAILED ✗"));
