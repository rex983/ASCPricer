/**
 * Final comprehensive review of snow engineering engine.
 * Re-traces ALL 10 configurations step by step, comparing
 * engine computations against spreadsheet-parsed data.
 */

import XLSX from "xlsx";

// ── Load and parse the spreadsheet ──
const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx");

function sheetToArray(ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
}
function num(v) {
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const n = parseFloat(v.replace(/[$,]/g, ""));
    return isNaN(n) ? 0 : n;
  }
  return 0;
}
function cleanHeader(v) {
  return String(v ?? "").trim().replace(/\s+/g, " ");
}

// ── Replicate readMatrix exactly ──
function readMatrix(data, opts = {}) {
  const { headerRow = 0, dataStartRow = 1, rowKeyCol = 0, dataStartCol = 1, transpose = false } = opts;
  const headers = data[headerRow] || [];
  const matrix = {};
  let endCol = opts.dataEndCol;
  if (endCol === undefined) {
    let last = dataStartCol;
    for (let c = dataStartCol; c < headers.length; c++) {
      const h = cleanHeader(headers[c]);
      if (h && h !== "0" && h !== "") last = c + 1;
    }
    endCol = last;
  }
  for (let r = dataStartRow; r < (opts.dataEndRow ?? data.length); r++) {
    const row = data[r];
    if (!row) break;
    const rowKey = cleanHeader(row[rowKeyCol]);
    if (!rowKey || rowKey === "0") continue;
    for (let c = dataStartCol; c < endCol; c++) {
      const colKey = cleanHeader(headers[c]);
      if (!colKey || colKey === "0") continue;
      const value = num(row[c]);
      if (transpose) {
        if (!matrix[rowKey]) matrix[rowKey] = {};
        matrix[rowKey][colKey] = value;
      } else {
        if (!matrix[colKey]) matrix[colKey] = {};
        matrix[colKey][rowKey] = value;
      }
    }
  }
  return matrix;
}

// ── Parse all snow sheets ──
function findSheet(name) {
  if (wb.Sheets[name]) return wb.Sheets[name];
  if (wb.Sheets[name + " "]) return wb.Sheets[name + " "];
  const trimmed = name.trim();
  for (const key of Object.keys(wb.Sheets)) {
    if (key.trim() === trimmed) return wb.Sheets[key];
  }
  return null;
}
function getSheet(name) {
  const ws = findSheet(name);
  if (!ws) throw new Error(`Sheet "${name}" not found. Available: ${wb.SheetNames.join(", ")}`);
  return ws;
}

console.log("Available sheets:", wb.SheetNames.join(", "));

// ── Parse Snow - Changers ──
function parseSnowChangers() {
  const ws = getSheet("Snow - Changers");
  const data = sheetToArray(ws);

  const windLoadBuckets = {};
  const heightClassification = {};
  const feetUsedByHeight = {};
  const pieTrussPrice = {};
  const trussPriceByWidthState = {};
  const channelPriceByState = {};
  const tubingPriceByState = {};

  // Wind load buckets (rows 0-1)
  if (data[0] && data[1]) {
    for (let c = 1; c < data[0].length; c++) {
      const inputMph = num(data[0][c]);
      const bucketMph = num(data[1][c]);
      if (inputMph >= 80 && inputMph <= 200 && bucketMph >= 100 && bucketMph <= 200) {
        windLoadBuckets[String(inputMph)] = bucketMph;
      }
    }
  }

  // Height classification (find S/M/T row)
  for (let r = 15; r < Math.min(35, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    let smtCount = 0;
    for (let c = 0; c < row.length; c++) {
      const v = cleanHeader(row[c]);
      if (v === "S" || v === "M" || v === "T") smtCount++;
    }
    if (smtCount >= 5) {
      const heightRow = data[r - 1];
      const feetRow = data[r + 1];
      if (heightRow) {
        for (let c = 0; c < row.length; c++) {
          const prefix = cleanHeader(row[c]);
          if (prefix !== "S" && prefix !== "M" && prefix !== "T") continue;
          const h = num(heightRow[c]);
          if (h >= 1 && h <= 30) {
            heightClassification[String(h)] = prefix === "S" ? 0 : prefix === "M" ? 1 : 2;
            if (feetRow) {
              feetUsedByHeight[String(h)] = num(feetRow[c]);
            }
          }
        }
      }
      break;
    }
  }

  // Per-state pricing
  let stateCodeRow = -1;
  let stateCodes = [];
  for (let r = 50; r < Math.min(75, data.length); r++) {
    const row = data[r];
    if (!row) continue;
    let codeCount = 0;
    const codes = [];
    for (let c = 1; c < row.length; c++) {
      const v = cleanHeader(row[c]);
      if (v.match(/^[A-Z]{2}$/)) { codeCount++; codes.push(v); }
      else codes.push("");
    }
    if (codeCount >= 5) { stateCodeRow = r; stateCodes = codes; break; }
  }

  if (stateCodeRow >= 0) {
    for (let r = stateCodeRow + 1; r < Math.min(stateCodeRow + 12, data.length); r++) {
      const row = data[r];
      if (!row) continue;
      const label = cleanHeader(row[0]).toLowerCase();
      if (label.includes("wide") || label.includes("truss")) {
        const widthMatch = label.match(/(\d+)/);
        if (widthMatch) {
          const w = widthMatch[1];
          for (let c = 1; c < row.length && c - 1 < stateCodes.length; c++) {
            const st = stateCodes[c - 1];
            if (!st) continue;
            const price = num(row[c]);
            if (price > 50 && price < 1000) {
              if (!trussPriceByWidthState[st]) trussPriceByWidthState[st] = {};
              trussPriceByWidthState[st][w] = price;
            }
          }
        }
      } else if (label.includes("pie")) {
        for (let c = 1; c < row.length && c - 1 < stateCodes.length; c++) {
          const st = stateCodes[c - 1]; if (!st) continue;
          const price = num(row[c]);
          if (price > 0 && price < 100) pieTrussPrice[st] = price;
        }
      } else if (label.includes("channel")) {
        for (let c = 1; c < row.length && c - 1 < stateCodes.length; c++) {
          const st = stateCodes[c - 1]; if (!st) continue;
          const price = num(row[c]);
          if (price >= 1 && price <= 10) channelPriceByState[st] = price;
        }
      } else if (label.includes("tubing") || label.includes("tube")) {
        for (let c = 1; c < row.length && c - 1 < stateCodes.length; c++) {
          const st = stateCodes[c - 1]; if (!st) continue;
          const price = num(row[c]);
          if (price >= 1 && price <= 10) tubingPriceByState[st] = price;
        }
      }
    }
  }

  return { windLoadBuckets, heightClassification, feetUsedByHeight, pieTrussPrice, trussPriceByWidthState, channelPriceByState, tubingPriceByState };
}

// ── Parse Truss Spacing ──
function parseTrussSpacing() {
  const ws = getSheet("Snow - Truss Spacing");
  const data = sheetToArray(ws);
  return readMatrix(data, { headerRow: 0, dataStartRow: 1, rowKeyCol: 0, dataStartCol: 1, transpose: true });
}

// ── Parse Trusses ──
function parseTrussCounts() {
  const ws = getSheet("Snow - Trusses");
  const data = sheetToArray(ws);
  return readMatrix(data, { headerRow: 0, dataStartRow: 1, rowKeyCol: 0, dataStartCol: 1 });
}

// ── Parse Hat Channels ──
function parseHatChannels() {
  const ws = getSheet("Snow - Hat Channels");
  const data = sheetToArray(ws);
  const headers = data[0] || [];
  const spacing = {};
  const originalCounts = {};

  let rightSectionStart = -1;
  for (let c = 8; c < headers.length; c++) {
    const h = cleanHeader(headers[c]);
    if (h.toLowerCase().includes("original")) { rightSectionStart = c; break; }
  }
  if (rightSectionStart < 0) {
    for (let c = 8; c < headers.length; c++) {
      const h = cleanHeader(headers[c]);
      if (["105","115","130","140","155","165","180","","0"].includes(h)) continue;
      let stateCount = 0;
      for (let r = 1; r < Math.min(10, data.length); r++) {
        const v = cleanHeader(data[r]?.[c] ?? "");
        if (v.match(/^[A-Z]{2}$/)) stateCount++;
      }
      if (stateCount >= 3) { rightSectionStart = c; break; }
    }
  }

  const spacingEndCol = rightSectionStart > 0 ? rightSectionStart : 8;
  for (let r = 1; r < data.length; r++) {
    const row = data[r]; if (!row) break;
    const rowKey = cleanHeader(row[0]); if (!rowKey) break;
    if (!spacing[rowKey]) spacing[rowKey] = {};
    for (let c = 1; c < spacingEndCol; c++) {
      const colKey = cleanHeader(headers[c]);
      if (!colKey || colKey === "0") continue;
      spacing[rowKey][colKey] = num(row[c]);
    }
  }

  if (rightSectionStart > 0) {
    let stateCol = rightSectionStart;
    let widthStartCol = rightSectionStart + 1;
    const firstHeader = cleanHeader(headers[rightSectionStart]);
    if (firstHeader.toLowerCase().includes("original") || !firstHeader.match(/^\d+$/)) {
      stateCol = rightSectionStart; widthStartCol = rightSectionStart + 1;
    }
    let widthHeaders = [];
    for (let c = widthStartCol; c < headers.length; c++) {
      const h = cleanHeader(headers[c]);
      if (h && num(h) >= 12 && num(h) <= 30) widthHeaders.push(h);
      else widthHeaders.push("");
    }
    for (let r = 1; r < data.length; r++) {
      const row = data[r]; if (!row) break;
      const state = cleanHeader(row[stateCol]);
      if (!state || !state.match(/^[A-Z]{2}$/)) continue;
      if (!originalCounts[state]) originalCounts[state] = {};
      for (let c = widthStartCol; c < row.length; c++) {
        const widthKey = widthHeaders[c - widthStartCol];
        if (!widthKey) continue;
        const count = num(row[c]);
        if (count > 0) originalCounts[state][widthKey] = count;
      }
    }
  }
  return { spacing, originalCounts };
}

// ── Parse Girts ──
function parseGirts() {
  const ws = getSheet("Snow - Girts");
  const data = sheetToArray(ws);
  const headers = data[0] || [];
  const spacing = {};
  const girtCountsByHeight = {};

  let rightSectionStart = -1;
  for (let c = 1; c < headers.length; c++) {
    const h = cleanHeader(headers[c]);
    if (h === "" || h === "0") {
      if (rightSectionStart < 0) {
        const next = cleanHeader(headers[c + 1]);
        if (next && !["105","115","130","140","155","165","180"].includes(next)) {
          rightSectionStart = c + 1; break;
        }
      }
    }
  }

  const spacingEndCol = rightSectionStart > 0 ? rightSectionStart : headers.length;
  for (let r = 1; r < data.length; r++) {
    const row = data[r]; if (!row) break;
    const rowKey = cleanHeader(row[0]); if (!rowKey) break;
    if (!spacing[rowKey]) spacing[rowKey] = {};
    for (let c = 1; c < spacingEndCol; c++) {
      const colKey = cleanHeader(headers[c]);
      if (!colKey || colKey === "0") continue;
      spacing[rowKey][colKey] = num(row[c]);
    }
  }

  if (rightSectionStart > 0) {
    for (let r = 1; r < data.length; r++) {
      const row = data[r]; if (!row) break;
      const heightKey = cleanHeader(row[rightSectionStart]);
      const count = num(row[rightSectionStart + 1]);
      if (!heightKey) continue;
      if (count > 0) girtCountsByHeight[heightKey] = count;
    }
  }
  return { spacing, girtCountsByHeight };
}

// ── Parse Verticals ──
function parseVerticals() {
  const ws = getSheet("Snow - Verticals");
  const data = sheetToArray(ws);
  const headers = data[0] || [];

  // Main spacing: verticalSpacing[windSpeed][height] → spacing, then we need [height][wind]
  // The engine uses lookupMatrix(verticalSpacing, String(height), String(wind))
  // readMatrix with default (no transpose) gives matrix[colKey][rowKey]
  // So headers[c] = heights (0,1,2,...20), row[0] = wind speeds (105,115,...)
  // Default: matrix[height][wind] — wait, let me check:
  // readMatrix default (transpose=false): matrix[colKey][rowKey] = value
  //   colKey = headers[c] = height values, rowKey = row[0] = wind speeds
  //   So matrix[height][windSpeed] = value ✓
  const spacing = readMatrix(data, { headerRow: 0, dataStartRow: 1, rowKeyCol: 0, dataStartCol: 1 });

  const originalCounts = {};
  for (let r = 7; r < Math.min(20, data.length); r++) {
    const row = data[r]; if (!row) continue;
    const label = cleanHeader(row[0]);
    if (label.toLowerCase().includes("original")) {
      const widthRow = row;
      const countRow = data[r + 1];
      if (countRow) {
        for (let c = 1; c < widthRow.length; c++) {
          const width = cleanHeader(widthRow[c]);
          if (width && num(width) >= 12) originalCounts[width] = num(countRow[c]);
        }
      }
      break;
    }
  }
  return { spacing, originalCounts };
}

// ── Parse everything ──
const snowChangers = parseSnowChangers();
const trussSpacing = parseTrussSpacing();
const trussCounts = parseTrussCounts();
const hatChannels = parseHatChannels();
const girts = parseGirts();
const verticals = parseVerticals();

// Build the matrices object matching StandardMatrices.snow
const matrices = {
  snow: {
    trussSpacing,
    trussCounts,
    hatChannelSpacing: hatChannels.spacing,
    hatChannelCounts: hatChannels.originalCounts,
    girtSpacing: girts.spacing,
    girtCountsByHeight: girts.girtCountsByHeight,
    verticalSpacing: verticals.spacing,
    verticalCounts: verticals.originalCounts,
    ...snowChangers,
  }
};

// ── Dump key parsed data for verification ──
console.log("\n=== PARSED DATA SUMMARY ===");
console.log("\nWind Load Buckets (sample):", JSON.stringify(snowChangers.windLoadBuckets));
console.log("\nHeight Classification:", JSON.stringify(snowChangers.heightClassification));
console.log("\nFeet Used By Height:", JSON.stringify(snowChangers.feetUsedByHeight));
console.log("\nPie Truss Price:", JSON.stringify(snowChangers.pieTrussPrice));
console.log("\nTruss Price By Width State:", JSON.stringify(snowChangers.trussPriceByWidthState));
console.log("\nChannel Price By State:", JSON.stringify(snowChangers.channelPriceByState));
console.log("\nTubing Price By State:", JSON.stringify(snowChangers.tubingPriceByState));

// Sample truss spacing keys
const tsKeys = Object.keys(trussSpacing);
console.log("\nTruss Spacing row keys (first 10):", tsKeys.slice(0, 10));
if (tsKeys.length > 0) {
  console.log("  Sample cols for", tsKeys[0], ":", Object.keys(trussSpacing[tsKeys[0]]).slice(0, 5));
}

// Sample truss counts keys
const tcKeys = Object.keys(trussCounts);
console.log("\nTruss Counts col keys (first 10):", tcKeys.slice(0, 10));

// Sample hat channel spacing
const hcKeys = Object.keys(hatChannels.spacing);
console.log("\nHat Channel Spacing row keys (first 10):", hcKeys.slice(0, 10));
if (hcKeys.length > 0) {
  console.log("  Sample cols for", hcKeys[0], ":", JSON.stringify(hatChannels.spacing[hcKeys[0]]));
}

// Hat channel original counts
console.log("\nHat Channel Original Counts:", JSON.stringify(hatChannels.originalCounts));

// Girt spacing
console.log("\nGirt Spacing:", JSON.stringify(girts.spacing));
console.log("\nGirt Counts By Height:", JSON.stringify(girts.girtCountsByHeight));

// Verticals
const vKeys = Object.keys(verticals.spacing);
console.log("\nVertical Spacing col keys:", vKeys.slice(0, 10));
if (vKeys.length > 0) {
  console.log("  Sample data for height '12':", JSON.stringify(verticals.spacing["12"]));
  console.log("  Sample data for height '16':", JSON.stringify(verticals.spacing["16"]));
}
console.log("\nVertical Original Counts:", JSON.stringify(verticals.originalCounts));


// ── Engine helpers (replicated from snow-engineering.ts) ──

const WIND_LOAD_CATEGORIES = [105, 115, 130, 140, 155, 165, 180];
const TRUSS_SPACING_BUCKETS = [36, 42, 48, 54, 60];
const ROOF_PITCH = 3 / 12;

function nearestBucket(value, buckets) {
  let nearest = buckets[0];
  for (const bucket of buckets) {
    if (bucket <= value) nearest = bucket;
    else break;
  }
  return nearest;
}

function lookupMatrix(matrix, rowKey, colKey) {
  return matrix[rowKey]?.[colKey] ?? 0;
}

function lookupValue(lookup, key) {
  return lookup[key] ?? 0;
}

function getHeightPrefix(height, classification) {
  const val = classification[String(height)];
  if (val !== undefined) {
    if (val === 0) return "S";
    if (val === 1) return "M";
    if (val === 2) return "T";
  }
  if (height <= 6) return "S";
  if (height <= 9) return "M";
  return "T";
}

function bucketWind(inputMph, buckets, categories) {
  const bucketed = buckets[String(inputMph)];
  if (bucketed && bucketed > 0) return bucketed;
  return nearestBucket(inputMph, categories);
}

function adjustTrussSpacingForHeight(rawSpacing, height) {
  let adjusted = rawSpacing;
  if (height >= 16) adjusted = rawSpacing - 12;
  else if (height >= 13) adjusted = rawSpacing - 6;
  if (adjusted <= 12) return 0;
  return adjusted;
}

function resolveEngineeringState(state, width, trussCounts) {
  const exactKey = `${width}-${state}`;
  if (trussCounts[exactKey]) return state;
  const prefix = `${width}-`;
  for (const key of Object.keys(trussCounts)) {
    if (key.startsWith(prefix)) return key.slice(prefix.length);
  }
  return state;
}

function getTrussPrice(width, state, trussPriceByWidthState) {
  const stateRow = trussPriceByWidthState[state];
  if (!stateRow) return 190;
  if (stateRow[String(width)] !== undefined) return stateRow[String(width)];
  for (const [rangeKey, price] of Object.entries(stateRow)) {
    const parts = rangeKey.split("-").map(Number);
    if (parts.length === 2 && width >= parts[0] && width <= parts[1]) return price;
  }
  return 190;
}

function getEffectiveFeetUsed(baseFeetUsed, width, height) {
  if (width >= 26 && height < 13) return baseFeetUsed * 2;
  return baseFeetUsed;
}

function getRoofRise(width, roofKey) {
  if (roofKey === "AFV" || roofKey === "AFH") return Math.ceil((width / 2) * ROOF_PITCH);
  return 0;
}

function getVerticalHeightMultiplier(height) {
  if (height >= 19) return 3.0;
  if (height >= 16) return 2.5;
  if (height >= 13) return 2.0;
  return 1.0;
}

function resolveOriginalGirts(height, girtCountsByHeight) {
  const exact = girtCountsByHeight[String(height)];
  if (exact > 0) return exact;
  for (const [key, count] of Object.entries(girtCountsByHeight)) {
    const parts = key.split("-").map(Number);
    if (parts.length === 2 && height >= parts[0] && height <= parts[1]) return count;
  }
  if (height <= 11) return 3;
  if (height <= 17) return 4;
  return 5;
}


// ── Test Configurations ──
const configs = [
  { id: 1, width: 22, length: 50, height: 12, roof: "AFV", orient: "horizontal", coverage: "fully_enclosed", sidesQty: 2, endsQty: 2, snow: "30GL", wind: 90, state: "AZ" },
  { id: 2, width: 30, length: 100, height: 15, roof: "AFV", orient: "horizontal", coverage: "fully_enclosed", sidesQty: 2, endsQty: 2, snow: "70GL", wind: 100, state: "AZ" },
  { id: 3, width: 24, length: 50, height: 12, roof: "AFV", orient: "horizontal", coverage: "fully_enclosed", sidesQty: 2, endsQty: 2, snow: "20LL", wind: 90, state: "AZ" },
  { id: 4, width: 24, length: 75, height: 14, roof: "AFV", orient: "vertical", coverage: "fully_enclosed", sidesQty: 2, endsQty: 2, snow: "50GL", wind: 130, state: "AZ" },
  { id: 5, width: 30, length: 100, height: 18, roof: "STD", orient: "vertical", coverage: "fully_enclosed", sidesQty: 2, endsQty: 2, snow: "90GL", wind: 155, state: "AZ" },
  { id: 6, width: 12, length: 25, height: 8, roof: "AFV", orient: "horizontal", coverage: "open", sidesQty: 0, endsQty: 0, snow: "30GL", wind: 105, state: "AZ" },
  { id: 7, width: 26, length: 50, height: 16, roof: "AFV", orient: "horizontal", coverage: "fully_enclosed", sidesQty: 2, endsQty: 2, snow: "70GL", wind: 140, state: "AZ" },
  { id: 8, width: 24, length: 100, height: 12, roof: "AFV", orient: "horizontal", coverage: "fully_enclosed", sidesQty: 2, endsQty: 2, snow: "40GL", wind: 115, state: "AZ" },
  { id: 9, width: 20, length: 50, height: 10, roof: "STD", orient: "horizontal", coverage: "open", sidesQty: 0, endsQty: 0, snow: "20LL", wind: 105, state: "AZ" },
  { id: 10, width: 28, length: 75, height: 20, roof: "AFV", orient: "vertical", coverage: "fully_enclosed", sidesQty: 2, endsQty: 2, snow: "80GL", wind: 165, state: "AZ" },
];

// ── Trace each config ──
const results = [];

for (const cfg of configs) {
  console.log(`\n${"=".repeat(80)}`);
  console.log(`CONFIG ${cfg.id}: ${cfg.width}x${cfg.length}x${cfg.height}, ${cfg.roof}, ${cfg.orient}, ${cfg.coverage}, ${cfg.sidesQty}S/${cfg.endsQty}E, ${cfg.snow}, ${cfg.wind}mph, ${cfg.state}`);
  console.log("=".repeat(80));

  const trace = {};

  // Step 1: Wind bucketing
  const bucketedWind = bucketWind(cfg.wind, matrices.snow.windLoadBuckets, WIND_LOAD_CATEGORIES);
  trace.bucketedWind = bucketedWind;
  console.log(`\n  Step 1 - Wind Bucketing: ${cfg.wind}mph → ${bucketedWind}mph`);
  console.log(`    Lookup windLoadBuckets["${cfg.wind}"]:`, matrices.snow.windLoadBuckets[String(cfg.wind)] ?? "(not found, using nearestBucket)");

  // Step 2: Height prefix
  const resolvedState = resolveEngineeringState(cfg.state, cfg.width, matrices.snow.trussCounts);
  const heightPrefix = getHeightPrefix(cfg.height, matrices.snow.heightClassification);
  trace.heightPrefix = heightPrefix;
  trace.resolvedState = resolvedState;
  console.log(`  Step 2 - Height Prefix: height=${cfg.height} → "${heightPrefix}" (classification value: ${matrices.snow.heightClassification[String(cfg.height)]})`);
  console.log(`           State resolved: "${cfg.state}" → "${resolvedState}"`);

  // Step 3: Snow code
  const snowCode = `${heightPrefix}-${cfg.snow}`;
  trace.snowCode = snowCode;
  console.log(`  Step 3 - Snow Code: "${snowCode}"`);

  // Step 4: E/O determination
  const isEnclosed = cfg.coverage !== "open" && cfg.sidesQty >= 2 && cfg.endsQty >= 2;
  trace.isEnclosed = isEnclosed;
  const enclosure = isEnclosed ? "E" : "O";
  console.log(`  Step 4 - E/O: coverage="${cfg.coverage}", sidesQty=${cfg.sidesQty}, endsQty=${cfg.endsQty} → "${enclosure}"`);

  // Step 5: Config key
  const configKey = `${enclosure}-${bucketedWind}-${cfg.width}-${cfg.roof}`;
  trace.configKey = configKey;
  console.log(`  Step 5 - Config Key: "${configKey}"`);

  // Step 6: Raw truss spacing lookup
  const rawTrussSpacing = lookupMatrix(matrices.snow.trussSpacing, snowCode, configKey);
  trace.rawTrussSpacing = rawTrussSpacing;
  console.log(`  Step 6 - Raw Truss Spacing: trussSpacing["${snowCode}"]["${configKey}"] = ${rawTrussSpacing}`);

  // Check if the row exists at all
  if (!matrices.snow.trussSpacing[snowCode]) {
    console.log(`    *** WARNING: Snow code row "${snowCode}" NOT FOUND in truss spacing matrix!`);
    // Show available rows that start similarly
    const matchingRows = Object.keys(matrices.snow.trussSpacing).filter(k => k.includes(cfg.snow));
    console.log(`    Available rows with "${cfg.snow}":`, matchingRows);
  } else if (matrices.snow.trussSpacing[snowCode][configKey] === undefined) {
    console.log(`    *** WARNING: Config key "${configKey}" NOT FOUND for snow code "${snowCode}"!`);
    // Show available cols
    const availCols = Object.keys(matrices.snow.trussSpacing[snowCode]).filter(k => k.includes(String(cfg.width)));
    console.log(`    Available cols with width=${cfg.width}:`, availCols);
  }

  // Step 7: F54 height adjustment
  const trussSpacingAdj = adjustTrussSpacingForHeight(rawTrussSpacing, cfg.height);
  trace.trussSpacingAdj = trussSpacingAdj;
  let adjNote = "";
  if (cfg.height >= 16) adjNote = `${rawTrussSpacing} - 12 = ${rawTrussSpacing - 12}`;
  else if (cfg.height >= 13) adjNote = `${rawTrussSpacing} - 6 = ${rawTrussSpacing - 6}`;
  else adjNote = `no adjustment (height ${cfg.height} ≤ 12)`;
  console.log(`  Step 7 - Height Adjustment (F54): ${adjNote} → effective spacing = ${trussSpacingAdj}`);

  // Step 8: Contact Engineering check
  const contactEng = trussSpacingAdj === 0 || trussSpacingAdj < 18;
  trace.contactEngineering = contactEng;
  console.log(`  Step 8 - Contact Engineering: spacing=${trussSpacingAdj}, check (===0 OR <18): ${contactEng ? "YES - CONTACT ENGINEERING" : "No"}`);

  if (contactEng) {
    trace.totalCost = -1;
    trace.trussCost = 0;
    trace.hcCost = 0;
    trace.girtCost = 0;
    trace.verticalCost = 0;
    results.push({ cfg, trace });
    console.log(`  >>> RESULT: Contact Engineering (return -1)`);
    continue;
  }

  // Step 9: Truss count and cost
  let totalCost = 0;
  const lengthInches = cfg.length * 12;
  const trussesNeeded = Math.ceil(lengthInches / trussSpacingAdj) + 1;
  const widthStateKey = `${cfg.width}-${resolvedState}`;
  const originalTrusses = lookupMatrix(matrices.snow.trussCounts, widthStateKey, String(cfg.length));
  const extraTrusses = Math.max(0, trussesNeeded - originalTrusses);

  console.log(`\n  Step 9 - Trusses:`);
  console.log(`    lengthInches = ${cfg.length} × 12 = ${lengthInches}`);
  console.log(`    trussesNeeded = ceil(${lengthInches} / ${trussSpacingAdj}) + 1 = ${Math.ceil(lengthInches / trussSpacingAdj)} + 1 = ${trussesNeeded}`);
  console.log(`    originalTrusses = trussCounts["${widthStateKey}"]["${cfg.length}"] = ${originalTrusses}`);
  console.log(`    extraTrusses = max(0, ${trussesNeeded} - ${originalTrusses}) = ${extraTrusses}`);

  let trussCost = 0;
  if (extraTrusses > 0) {
    const baseTrussPrice = getTrussPrice(cfg.width, resolvedState, matrices.snow.trussPriceByWidthState);
    const baseFeetUsed = lookupValue(matrices.snow.feetUsedByHeight, String(cfg.height));
    const feetUsed = getEffectiveFeetUsed(baseFeetUsed, cfg.width, cfg.height);
    const piePricePerFt = lookupValue(matrices.snow.pieTrussPrice, resolvedState) || 15;
    const legSurcharge = feetUsed * piePricePerFt;
    trussCost = extraTrusses * (baseTrussPrice + legSurcharge);

    console.log(`    baseTrussPrice = getTrussPrice(${cfg.width}, "${resolvedState}") = ${baseTrussPrice}`);
    console.log(`    baseFeetUsed = feetUsedByHeight["${cfg.height}"] = ${baseFeetUsed}`);
    console.log(`    feetUsed = ${cfg.width >= 26 && cfg.height < 13 ? `${baseFeetUsed}×2=${feetUsed} (width≥26, height<13)` : `${feetUsed} (no doubling)`}`);
    console.log(`    piePricePerFt = pieTrussPrice["${resolvedState}"] = ${piePricePerFt}`);
    console.log(`    legSurcharge = ${feetUsed} × ${piePricePerFt} = ${legSurcharge}`);
    console.log(`    trussCost = ${extraTrusses} × (${baseTrussPrice} + ${legSurcharge}) = ${trussCost}`);
  } else {
    console.log(`    No extra trusses needed → trussCost = 0`);
  }
  totalCost += trussCost;
  trace.trussCost = trussCost;

  // Step 10: Hat Channels
  const isAFV = cfg.roof === "AFV" || cfg.roof === "AFH";
  const actualTrussSpacing = trussSpacingAdj > 0 ? trussSpacingAdj : 60;
  const bucketedTrussSpacing = nearestBucket(actualTrussSpacing, TRUSS_SPACING_BUCKETS);
  const hcRowKey = `${bucketedTrussSpacing}-${cfg.snow}`;
  const hatChannelSpacingVal = lookupMatrix(matrices.snow.hatChannelSpacing, hcRowKey, String(bucketedWind));

  console.log(`\n  Step 10 - Hat Channels:`);
  console.log(`    isAFV = ${isAFV} (roof="${cfg.roof}")`);
  console.log(`    actualTrussSpacing = ${actualTrussSpacing}`);
  console.log(`    bucketedTrussSpacing = nearestBucket(${actualTrussSpacing}, [36,42,48,54,60]) = ${bucketedTrussSpacing}`);
  console.log(`    hcRowKey = "${hcRowKey}"`);
  console.log(`    hatChannelSpacing = hatChannelSpacing["${hcRowKey}"]["${bucketedWind}"] = ${hatChannelSpacingVal}`);

  let hcCost = 0;
  if (isAFV && hatChannelSpacingVal > 0) {
    const barSize = (cfg.width + 2) / 2;
    const barInches = barSize * 12;
    const channelsPerSide = Math.ceil(barInches / hatChannelSpacingVal) + 1;
    const totalChannels = channelsPerSide * 2;

    let originalChannels = lookupMatrix(matrices.snow.hatChannelCounts, resolvedState, String(cfg.width));
    if (originalChannels === 0) {
      originalChannels = lookupMatrix(matrices.snow.hatChannelCounts, String(cfg.width), resolvedState);
    }
    const extraChannels = Math.max(0, totalChannels - originalChannels);

    console.log(`    barSize = (${cfg.width} + 2) / 2 = ${barSize}`);
    console.log(`    barInches = ${barSize} × 12 = ${barInches}`);
    console.log(`    channelsPerSide = ceil(${barInches} / ${hatChannelSpacingVal}) + 1 = ${channelsPerSide}`);
    console.log(`    totalChannels = ${channelsPerSide} × 2 = ${totalChannels}`);
    console.log(`    originalChannels = hatChannelCounts["${resolvedState}"]["${cfg.width}"] = ${originalChannels}`);
    console.log(`    extraChannels = max(0, ${totalChannels} - ${originalChannels}) = ${extraChannels}`);

    if (extraChannels > 0) {
      const channelPricePerFt = lookupValue(matrices.snow.channelPriceByState, resolvedState) || 2;
      const channelLength = cfg.length + 1;
      hcCost = extraChannels * channelPricePerFt * channelLength;
      console.log(`    channelPricePerFt = ${channelPricePerFt}`);
      console.log(`    channelLength = ${cfg.length} + 1 = ${channelLength}`);
      console.log(`    hcCost = ${extraChannels} × ${channelPricePerFt} × ${channelLength} = ${hcCost}`);
    } else {
      console.log(`    No extra hat channels needed → hcCost = 0`);
    }
  } else {
    console.log(`    ${!isAFV ? "Not AFV roof → " : "HC spacing=0 → "}hcCost = 0`);
  }
  totalCost += hcCost;
  trace.hcCost = hcCost;

  // Step 11: Girts
  const hasVerticalPanels = cfg.orient === "vertical";
  const girtsNeeded = isEnclosed && hasVerticalPanels;

  console.log(`\n  Step 11 - Girts:`);
  console.log(`    isEnclosed = ${isEnclosed}, hasVerticalPanels = ${hasVerticalPanels}`);
  console.log(`    girtsNeeded (double gate) = ${isEnclosed} AND ${hasVerticalPanels} = ${girtsNeeded}`);

  let girtCost = 0;
  if (girtsNeeded) {
    const girtSpacingVal = lookupMatrix(matrices.snow.girtSpacing, String(bucketedTrussSpacing), String(bucketedWind));
    console.log(`    girtSpacing = girtSpacing["${bucketedTrussSpacing}"]["${bucketedWind}"] = ${girtSpacingVal}`);

    if (girtSpacingVal > 0) {
      const heightInches = cfg.height * 12;
      const girtsRequired = Math.ceil(heightInches / girtSpacingVal) + 1;
      const originalGirts = resolveOriginalGirts(cfg.height, matrices.snow.girtCountsByHeight);
      const extraGirts = Math.max(0, girtsRequired - originalGirts);

      console.log(`    heightInches = ${cfg.height} × 12 = ${heightInches}`);
      console.log(`    girtsRequired = ceil(${heightInches} / ${girtSpacingVal}) + 1 = ${girtsRequired}`);
      console.log(`    originalGirts = resolveOriginalGirts(${cfg.height}) = ${originalGirts}`);
      console.log(`    extraGirts = max(0, ${girtsRequired} - ${originalGirts}) = ${extraGirts}`);

      if (extraGirts > 0) {
        const tubingPrice = lookupValue(matrices.snow.tubingPriceByState, resolvedState) || 3;
        let perimeter = 0;
        // Sides: only if coverage != open AND sidesOrientation == vertical
        if (cfg.coverage !== "open" && cfg.orient === "vertical") {
          perimeter += cfg.sidesQty * cfg.length;
        }
        // Ends: only if endsQty > 0 AND endsOrientation == vertical
        // NOTE: In the engine, config.endsOrientation is checked, but in our configs orient applies to sides
        // For simplicity, configs with vertical orient have both sides and ends vertical
        if (cfg.endsQty > 0 && cfg.orient === "vertical") {
          perimeter += cfg.endsQty * cfg.width;
        }
        girtCost = extraGirts * tubingPrice * perimeter;

        console.log(`    tubingPrice = ${tubingPrice}`);
        console.log(`    perimeter = ${cfg.coverage !== "open" && cfg.orient === "vertical" ? `${cfg.sidesQty}×${cfg.length}` : "0"} + ${cfg.endsQty > 0 && cfg.orient === "vertical" ? `${cfg.endsQty}×${cfg.width}` : "0"} = ${perimeter}`);
        console.log(`    girtCost = ${extraGirts} × ${tubingPrice} × ${perimeter} = ${girtCost}`);
      } else {
        console.log(`    No extra girts needed → girtCost = 0`);
      }
    } else {
      console.log(`    girtSpacing = 0 → girtCost = 0`);
    }
  } else {
    console.log(`    Double gate not met → girtCost = 0`);
  }
  totalCost += girtCost;
  trace.girtCost = girtCost;

  // Step 12: Verticals
  const verticalSpacingVal = lookupMatrix(matrices.snow.verticalSpacing, String(cfg.height), String(bucketedWind));

  console.log(`\n  Step 12 - Verticals:`);
  console.log(`    verticalSpacing = verticalSpacing["${cfg.height}"]["${bucketedWind}"] = ${verticalSpacingVal}`);

  let verticalCost = 0;
  if (verticalSpacingVal > 0) {
    const widthInches = cfg.width * 12;
    const verticalsNeeded = Math.ceil(widthInches / verticalSpacingVal) + 1;
    const originalVerticals = lookupValue(matrices.snow.verticalCounts, String(cfg.width));
    const extraVerticals = Math.max(0, verticalsNeeded - originalVerticals);

    console.log(`    widthInches = ${cfg.width} × 12 = ${widthInches}`);
    console.log(`    verticalsNeeded = ceil(${widthInches} / ${verticalSpacingVal}) + 1 = ${verticalsNeeded}`);
    console.log(`    originalVerticals = verticalCounts["${cfg.width}"] = ${originalVerticals}`);
    console.log(`    extraVerticals = max(0, ${verticalsNeeded} - ${originalVerticals}) = ${extraVerticals}`);

    if (extraVerticals > 0 && cfg.endsQty > 0) {
      const tubingPrice = lookupValue(matrices.snow.tubingPriceByState, resolvedState) || 3;
      const roofRise = getRoofRise(cfg.width, cfg.roof);
      const peakHeight = cfg.height + roofRise;
      const heightMult = getVerticalHeightMultiplier(cfg.height);
      const baseVertCost = extraVerticals * cfg.endsQty * tubingPrice * peakHeight;
      verticalCost = baseVertCost * heightMult;

      console.log(`    tubingPrice = ${tubingPrice}`);
      console.log(`    roofRise = getRoofRise(${cfg.width}, "${cfg.roof}") = ${roofRise}`);
      console.log(`    peakHeight = ${cfg.height} + ${roofRise} = ${peakHeight}`);
      console.log(`    heightMult = getVerticalHeightMultiplier(${cfg.height}) = ${heightMult}`);
      console.log(`    baseVertCost = ${extraVerticals} × ${cfg.endsQty} × ${tubingPrice} × ${peakHeight} = ${baseVertCost}`);
      console.log(`    verticalCost = ${baseVertCost} × ${heightMult} = ${verticalCost}`);
    } else {
      console.log(`    ${extraVerticals <= 0 ? "No extra verticals" : "endsQty=0"} → verticalCost = 0`);
    }
  } else {
    console.log(`    verticalSpacing = 0 → verticalCost = 0`);
  }
  totalCost += verticalCost;
  trace.verticalCost = verticalCost;

  // Step 13: Total
  trace.totalCost = Math.round(totalCost);
  console.log(`\n  Step 13 - TOTAL:`);
  console.log(`    trussCost    = ${trussCost}`);
  console.log(`    hcCost       = ${hcCost}`);
  console.log(`    girtCost     = ${girtCost}`);
  console.log(`    verticalCost = ${verticalCost}`);
  console.log(`    TOTAL        = round(${totalCost}) = ${trace.totalCost}`);

  results.push({ cfg, trace });
}

// ── Summary Table ──
console.log(`\n\n${"=".repeat(120)}`);
console.log("SUMMARY TABLE");
console.log("=".repeat(120));
console.log(
  "Config".padEnd(8) +
  "Size".padEnd(14) +
  "Roof".padEnd(6) +
  "Snow".padEnd(8) +
  "Wind".padEnd(8) +
  "BktWind".padEnd(9) +
  "HtPfx".padEnd(7) +
  "E/O".padEnd(5) +
  "ConfigKey".padEnd(22) +
  "RawSpc".padEnd(8) +
  "AdjSpc".padEnd(8) +
  "Truss$".padEnd(10) +
  "HC$".padEnd(10) +
  "Girt$".padEnd(10) +
  "Vert$".padEnd(10) +
  "TOTAL"
);
console.log("-".repeat(120));

for (const { cfg, trace } of results) {
  const size = `${cfg.width}x${cfg.length}x${cfg.height}`;
  const total = trace.contactEngineering ? "CONTACT ENG" : `$${trace.totalCost.toLocaleString()}`;
  console.log(
    String(cfg.id).padEnd(8) +
    size.padEnd(14) +
    cfg.roof.padEnd(6) +
    cfg.snow.padEnd(8) +
    String(cfg.wind).padEnd(8) +
    String(trace.bucketedWind).padEnd(9) +
    trace.heightPrefix.padEnd(7) +
    (trace.isEnclosed ? "E" : "O").padEnd(5) +
    trace.configKey.padEnd(22) +
    String(trace.rawTrussSpacing).padEnd(8) +
    String(trace.trussSpacingAdj).padEnd(8) +
    (trace.contactEngineering ? "---" : `$${trace.trussCost}`).padEnd(10) +
    (trace.contactEngineering ? "---" : `$${trace.hcCost}`).padEnd(10) +
    (trace.contactEngineering ? "---" : `$${trace.girtCost}`).padEnd(10) +
    (trace.contactEngineering ? "---" : `$${trace.verticalCost}`).padEnd(10) +
    total
  );
}

// ── Edge Case Verification ──
console.log(`\n\n${"=".repeat(80)}`);
console.log("EDGE CASE VERIFICATION");
console.log("=".repeat(80));

// Config 5: STD roof, 90GL, 155mph → Contact Engineering
{
  const r = results.find(r => r.cfg.id === 5);
  const t = r.trace;
  console.log(`\n[Config 5] STD, 90GL, 155mph, height=18:`);
  console.log(`  Snow code: "${t.snowCode}"`);
  console.log(`  Config key: "${t.configKey}"`);
  console.log(`  Raw truss spacing = ${t.rawTrussSpacing}`);
  console.log(`  Height adj: 18 >= 16 → -12, so ${t.rawTrussSpacing} - 12 = ${t.rawTrussSpacing - 12}`);
  console.log(`  Adjusted spacing = ${t.trussSpacingAdj}`);
  console.log(`  Contact Engineering? ${t.contactEngineering}`);
  if (t.rawTrussSpacing === 0) {
    console.log(`  VERIFIED: Raw spacing=0, lookup not found → Contact Engineering`);
  } else if (t.trussSpacingAdj === 0 || t.trussSpacingAdj < 18) {
    console.log(`  VERIFIED: Adjusted spacing ${t.trussSpacingAdj} triggers Contact Engineering`);
  } else {
    console.log(`  *** UNEXPECTED: Expected Contact Engineering but got spacing=${t.trussSpacingAdj}`);
  }
}

// Config 7: height 16, F54 = -12
{
  const r = results.find(r => r.cfg.id === 7);
  const t = r.trace;
  console.log(`\n[Config 7] AFV, 70GL, 140mph, height=16:`);
  console.log(`  Raw truss spacing = ${t.rawTrussSpacing}`);
  console.log(`  Height 16 >= 16 → F54 reduction = -12 (not -6)`);
  console.log(`  Adjusted = ${t.rawTrussSpacing} - 12 = ${t.rawTrussSpacing - 12}, engine result = ${t.trussSpacingAdj}`);
  if (t.trussSpacingAdj === t.rawTrussSpacing - 12 || (t.rawTrussSpacing - 12 <= 12 && t.trussSpacingAdj === 0)) {
    console.log(`  VERIFIED: F54 reduction is -12 for height 16`);
  } else {
    console.log(`  *** BUG: Expected -12 reduction but got different result`);
  }

  // Check HC bucket uses adjusted spacing
  const actualTS = t.trussSpacingAdj > 0 ? t.trussSpacingAdj : 60;
  const bucketedTS = nearestBucket(actualTS, TRUSS_SPACING_BUCKETS);
  console.log(`  HC bucket: actualTrussSpacing=${actualTS}, bucketedTrussSpacing=${bucketedTS}`);
  console.log(`  HC uses ADJUSTED spacing for bucketing: ${t.trussSpacingAdj > 0 ? "YES (uses " + t.trussSpacingAdj + ")" : "N/A (Contact Eng)"}`);
}

// Config 10: height 20, check if spacing reduces to ≤12
{
  const r = results.find(r => r.cfg.id === 10);
  const t = r.trace;
  console.log(`\n[Config 10] AFV, 80GL, 165mph, height=20:`);
  console.log(`  Raw truss spacing = ${t.rawTrussSpacing}`);
  console.log(`  Height 20 >= 16 → F54 reduction = -12`);
  console.log(`  Adjusted = ${t.rawTrussSpacing} - 12 = ${t.rawTrussSpacing - 12}, engine result = ${t.trussSpacingAdj}`);
  if (t.contactEngineering) {
    console.log(`  VERIFIED: Spacing reduced to ${t.trussSpacingAdj} → Contact Engineering`);
  } else {
    console.log(`  Result: NOT Contact Engineering, adjusted spacing = ${t.trussSpacingAdj}`);
  }
}

// Configs 5, 9: STD roof → HC extras = 0
{
  console.log(`\n[Configs 5, 9] STD roof → HC extras should be 0:`);
  for (const id of [5, 9]) {
    const r = results.find(r => r.cfg.id === id);
    const t = r.trace;
    console.log(`  Config ${id} (${r.cfg.roof}): hcCost = ${t.hcCost}${t.contactEngineering ? " (Contact Eng)" : ""}`);
    if (r.cfg.roof === "STD" && (t.hcCost === 0 || t.contactEngineering)) {
      console.log(`    VERIFIED: STD roof correctly gets 0 HC extras`);
    } else {
      console.log(`    *** BUG: STD roof should get 0 HC extras but got ${t.hcCost}`);
    }
  }
}

// ── Analysis: F52 vs F54 discrepancy for HC bucketing ──
console.log(`\n\n${"=".repeat(80)}`);
console.log("ANALYSIS: F52 vs F54 for HC Bucketing (heights 16+)");
console.log("=".repeat(80));
console.log(`\nPer the analysis doc:`);
console.log(`  F52 = raw - I106 (where I106=6 for heights 13+), used for HC bucketing`);
console.log(`  F54 = raw - 6 (heights 13-15), raw - 12 (heights 16-20), used for truss count`);
console.log(`  For heights 16+: F52 = raw - 6, but F54 = raw - 12`);
console.log(`  The ENGINE uses F54 (trussSpacingAdj) for HC bucketing`);
console.log(`  This is a known minor discrepancy vs spreadsheet (which uses F52)`);
console.log(`  In practice, this rarely causes different HC results because:`);
console.log(`    - Both values often round to the same bucket`);
console.log(`    - Heights 16+ often result in Contact Engineering anyway`);

// Check configs 7 and 10 for this
for (const id of [7, 10]) {
  const r = results.find(r => r.cfg.id === id);
  const t = r.trace;
  if (!t.contactEngineering && t.rawTrussSpacing > 0) {
    const f52 = t.rawTrussSpacing - 6;  // what spreadsheet would use
    const f54 = t.rawTrussSpacing - 12; // what engine uses
    const bucketF52 = nearestBucket(f52, TRUSS_SPACING_BUCKETS);
    const bucketF54 = nearestBucket(f54 > 12 ? f54 : 60, TRUSS_SPACING_BUCKETS);
    console.log(`\n  Config ${id} (height=${r.cfg.height}):`);
    console.log(`    raw=${t.rawTrussSpacing}, F52=${f52} (bucket=${bucketF52}), F54=${f54} (bucket=${bucketF54})`);
    if (bucketF52 !== bucketF54) {
      console.log(`    *** DISCREPANCY: HC bucket would differ (F52 bucket=${bucketF52} vs F54 bucket=${bucketF54})`);
    } else {
      console.log(`    OK: Same HC bucket either way`);
    }
  } else {
    console.log(`\n  Config ${id}: ${t.contactEngineering ? "Contact Engineering (N/A)" : "raw=0 (N/A)"}`);
  }
}

console.log(`\n\nFinal review complete.`);
