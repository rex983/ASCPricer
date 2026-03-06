import * as XLSX from "xlsx";
import { parseSpreadsheet } from "../src/lib/excel/parser";
import * as fs from "fs";

async function test() {
  // Test Standard Spreadsheet
  console.log("=== TESTING STANDARD SPREADSHEET ===\n");
  const stdBuffer = fs.readFileSync("C:/Users/Redir/spreadsheet.xlsx");
  const stdResult = parseSpreadsheet(stdBuffer);

  console.log("Detection:", stdResult.detection);
  console.log("Validation:", stdResult.validation);
  console.log("Type:", stdResult.matrices.type);

  if (stdResult.matrices.type === "standard") {
    const m = stdResult.matrices;
    console.log("\nBase Price keys:", Object.keys(m.basePrice).slice(0, 5), "...");
    const firstKey = Object.keys(m.basePrice)[0];
    if (firstKey) {
      console.log(`  ${firstKey} entries:`, Object.entries(m.basePrice[firstKey]).slice(0, 5));
    }

    console.log("\nRoof Style keys:", Object.keys(m.roofStyle).slice(0, 5), "...");
    console.log("Legs small keys:", Object.keys(m.legs.small).slice(0, 5));
    console.log("Legs large keys:", Object.keys(m.legs.large).slice(0, 5));
    console.log("Sides keys:", Object.keys(m.sides).slice(0, 5), "...");
    console.log("Ends keys:", Object.keys(m.ends).slice(0, 5), "...");
    console.log("Walk-in Doors:", m.accessories.walkInDoors);
    console.log("Windows:", m.accessories.windows);
    console.log("Labor keys:", Object.keys(m.laborEquipment).slice(0, 5), "...");
    console.log("Plans keys:", Object.keys(m.plans).slice(0, 5), "...");
    console.log("Changers widthBuckets:", m.changers.widthBuckets);
    console.log("Changers lengthBuckets:", m.changers.lengthBuckets);
    console.log("Snow trussSpacing keys:", Object.keys(m.snow.trussSpacing).slice(0, 5), "...");

    // Spot check: 24-14G at length 50
    console.log("\n--- SPOT CHECK ---");
    console.log("Base Price 24-14G @ 50:", m.basePrice["24-14G"]?.["50"]);
    console.log("Roof Style AFV-24 @ 50:", m.roofStyle["AFV-24"]?.["50"]);
  }

  // Test Widespan Spreadsheet
  console.log("\n\n=== TESTING WIDESPAN SPREADSHEET ===\n");
  const wsBuffer = fs.readFileSync("C:/Users/Redir/spreadsheet-widespan.xlsx");
  const wsResult = parseSpreadsheet(wsBuffer);

  console.log("Detection:", wsResult.detection);
  console.log("Validation:", wsResult.validation);
  console.log("Type:", wsResult.matrices.type);

  if (wsResult.matrices.type === "widespan") {
    const m = wsResult.matrices;
    console.log("\nBase Price keys:", Object.keys(m.basePrice).slice(0, 5), "...");
    const firstKey = Object.keys(m.basePrice)[0];
    if (firstKey) {
      console.log(`  ${firstKey} entries:`, Object.entries(m.basePrice[firstKey]).slice(0, 5));
    }

    console.log("\nLegs keys:", Object.keys(m.legs).slice(0, 5), "...");
    console.log("Sides keys:", Object.keys(m.sides).slice(0, 5), "...");
    console.log("Ends keys:", Object.keys(m.ends).slice(0, 5), "...");
    console.log("Walk-in Doors:", m.accessories.walkInDoors);
    console.log("Windows:", m.accessories.windows);
    console.log("Roll-up Doors:", Object.entries(m.accessories.rollUpDoors).slice(0, 5), "...");
    console.log("Wainscot sides:", Object.entries(m.wainscot.sides).slice(0, 5), "...");
    console.log("Wainscot ends:", m.wainscot.ends);
    console.log("Equipment keys:", Object.keys(m.laborEquipment).slice(0, 5), "...");
    console.log("Plans keys:", Object.keys(m.plans).slice(0, 5), "...");
    console.log("Changers widthBuckets:", m.changers.widthBuckets);
    console.log("Changers lengthBuckets:", Object.entries(m.changers.lengthBuckets).slice(0, 10), "...");

    // Spot check: 60-12G at length 50
    console.log("\n--- SPOT CHECK ---");
    console.log("Base Price 60-12G @ 50:", m.basePrice["60-12G"]?.["50"]);
    console.log("Legs height 14 @ length 20:", m.legs["14"]?.["20"]);
  }
}

test().catch(console.error);
