import XLSX from "xlsx";
import { parseStandardWorkbook } from "./src/lib/excel/parser-standard.js";
import { calculateStandardSnowEngineering } from "./src/lib/pricing/snow-engineering.js";

const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx");
console.log("Sheets:", wb.SheetNames.map(s => JSON.stringify(s)).join(", "));

const matrices = parseStandardWorkbook(wb);
const s = matrices.snow;

console.log("\n--- Lookups ---");
console.log("windLoadBuckets[90]:", s.windLoadBuckets["90"]);
console.log("heightClass[10]:", s.heightClassification["10"]);
console.log("feetUsed[10]:", s.feetUsedByHeight["10"]);
console.log("pieTruss[AZ]:", s.pieTrussPrice["AZ"]);
console.log("trussSpacing T-20LL/E-105-24-AFV:", s.trussSpacing["T-20LL"]?.["E-105-24-AFV"]);
console.log("trussCounts 24-AZ/50:", s.trussCounts["24-AZ"]?.["50"]);
console.log("HC spacing 60-20LL/105:", s.hatChannelSpacing["60-20LL"]?.["105"]);
console.log("HC counts AZ:", JSON.stringify(s.hatChannelCounts["AZ"]));
console.log("girtSpacing 60/105:", s.girtSpacing["60"]?.["105"]);
console.log("girtCounts[10]:", s.girtCountsByHeight["10"]);
console.log("vertSpacing 10/105:", s.verticalSpacing["10"]?.["105"]);
console.log("vertCounts[24]:", s.verticalCounts["24"]);
console.log("trussPrice AZ:", JSON.stringify(s.trussPriceByWidthState["AZ"]));
console.log("channelPrice AZ:", s.channelPriceByState["AZ"]);
console.log("tubingPrice AZ:", s.tubingPriceByState["AZ"]);

const cost = calculateStandardSnowEngineering(
  { width: 24, length: 50, height: 10, gauge: 14, roofStyle: "a_frame_vertical",
    sidesCoverage: "fully_enclosed", sidesOrientation: "vertical", sidesQty: 2,
    endType: "enclosed", endsOrientation: "vertical", endsQty: 2,
    walkInDoorQty: 0, windowQty: 0, rollUpEndQty: 0, rollUpSideQty: 0,
    insulationType: "none", insulationQty: 0, snowLoad: "20LL", windRating: 90,
    diagonalBracing: false, state: "AZ", taxRate: 0 } as any,
  matrices, { width: 24, length: 50, roofKey: "AFV" }
);
console.log("\nSnow Engineering Cost: $" + cost);
