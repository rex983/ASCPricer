import XLSX from "xlsx";

const wb = XLSX.readFile("C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx");

function sheetToArray(ws) {
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
  const rows = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      row.push(cell ? (cell.v !== undefined ? cell.v : "") : "");
    }
    rows.push(row);
  }
  return rows;
}
function num(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function clean(v) { return String(v ?? "").trim(); }
function findSheet(name) {
  for (const key of Object.keys(wb.Sheets)) {
    if (key.trim() === name.trim()) return wb.Sheets[key];
  }
  return null;
}

// ── Replicate readMatrix from utils ──
function readMatrix(data, opts) {
  const { headerRow, dataStartRow, rowKeyCol, dataStartCol, transpose } = opts;
  const headers = data[headerRow] || [];
  const matrix = {};

  for (let r = dataStartRow; r < data.length; r++) {
    const row = data[r];
    if (!row) break;
    const rowKey = clean(row[rowKeyCol]);
    if (!rowKey) break;

    for (let c = dataStartCol; c < headers.length; c++) {
      const colKey = clean(headers[c]);
      if (!colKey || colKey === "0") continue;
      const val = num(row[c]);
      if (val === 0) continue;

      if (transpose) {
        if (!matrix[rowKey]) matrix[rowKey] = {};
        matrix[rowKey][colKey] = val;
      } else {
        if (!matrix[colKey]) matrix[colKey] = {};
        matrix[colKey][rowKey] = val;
      }
    }
  }
  return matrix;
}

// ── Parse Snow - Truss Spacing (transposed) ──
const tsSheet = findSheet("Snow - Truss Spacing");
const tsData = sheetToArray(tsSheet);
const trussSpacing = readMatrix(tsData, { headerRow: 0, dataStartRow: 1, rowKeyCol: 0, dataStartCol: 1, transpose: true });

// ── Parse Snow - Trusses (not transposed) ──
const trSheet = findSheet("Snow - Trusses");
const trData = sheetToArray(trSheet);
const trussCounts = readMatrix(trData, { headerRow: 0, dataStartRow: 1, rowKeyCol: 0, dataStartCol: 1 });

// ── Parse Snow - Hat Channels ──
const hcSheet = findSheet("Snow - Hat Channels");
const hcData = sheetToArray(hcSheet);
const hcHeaders = hcData[0] || [];

// Find "Original" section
let rightStart = -1;
for (let c = 8; c < hcHeaders.length; c++) {
  if (clean(hcHeaders[c]).toLowerCase().includes("original")) {
    rightStart = c;
    break;
  }
}

// Left: HC spacing
const hcSpacing = {};
const spacingEnd = rightStart > 0 ? rightStart : 8;
for (let r = 1; r < hcData.length; r++) {
  const row = hcData[r];
  if (!row) break;
  const rowKey = clean(row[0]);
  if (!rowKey) break;
  if (!hcSpacing[rowKey]) hcSpacing[rowKey] = {};
  for (let c = 1; c < spacingEnd; c++) {
    const colKey = clean(hcHeaders[c]);
    if (!colKey || colKey === "0") continue;
    hcSpacing[rowKey][colKey] = num(row[c]);
  }
}

// Right: HC original counts
const hcCounts = {};
if (rightStart > 0) {
  const stateCol = rightStart;
  const widthStartCol = rightStart + 1;
  const widthHeaders = [];
  for (let c = widthStartCol; c < hcHeaders.length; c++) {
    const h = clean(hcHeaders[c]);
    widthHeaders.push(num(h) >= 12 && num(h) <= 30 ? h : "");
  }
  for (let r = 1; r < hcData.length; r++) {
    const row = hcData[r];
    if (!row) break;
    const state = clean(row[stateCol]);
    if (!state.match(/^[A-Z]{2}$/)) continue;
    if (!hcCounts[state]) hcCounts[state] = {};
    for (let c = widthStartCol; c < row.length; c++) {
      const wk = widthHeaders[c - widthStartCol];
      if (!wk) continue;
      const count = num(row[c]);
      if (count > 0) hcCounts[state][wk] = count;
    }
  }
}

// ── Parse Snow - Girts ──
const gSheet = findSheet("Snow - Girts");
const gData = sheetToArray(gSheet);
const gHeaders = gData[0] || [];

// Find right section
let gRightStart = -1;
for (let c = 1; c < gHeaders.length; c++) {
  const h = clean(gHeaders[c]);
  if (h === "" || h === "0") {
    const next = clean(gHeaders[c + 1]);
    if (next && !["105", "115", "130", "140", "155", "165", "180"].includes(next)) {
      gRightStart = c + 1;
      break;
    }
  }
}

const girtSpacing = {};
const girtEnd = gRightStart > 0 ? gRightStart : gHeaders.length;
for (let r = 1; r < gData.length; r++) {
  const row = gData[r];
  if (!row) break;
  const rowKey = clean(row[0]);
  if (!rowKey) break;
  if (!girtSpacing[rowKey]) girtSpacing[rowKey] = {};
  for (let c = 1; c < girtEnd; c++) {
    const colKey = clean(gHeaders[c]);
    if (!colKey || colKey === "0") continue;
    girtSpacing[rowKey][colKey] = num(row[c]);
  }
}

const girtCountsByHeight = {};
if (gRightStart > 0) {
  for (let r = 1; r < gData.length; r++) {
    const row = gData[r];
    if (!row) break;
    const heightKey = clean(row[gRightStart]);
    const count = num(row[gRightStart + 1]);
    if (!heightKey) continue;
    if (count > 0) girtCountsByHeight[heightKey] = count;
  }
}

// ── Parse Snow - Verticals ──
const vSheet = findSheet("Snow - Verticals");
const vData = sheetToArray(vSheet);
const vHeaders = vData[0] || [];

// Main spacing (not transposed: matrix[colHeader][rowKey])
const vertSpacing = readMatrix(vData, { headerRow: 0, dataStartRow: 1, rowKeyCol: 0, dataStartCol: 1 });

// Original counts
const vertCounts = {};
for (let r = 7; r < Math.min(20, vData.length); r++) {
  const label = clean(vData[r]?.[0]);
  if (label.toLowerCase().includes("original")) {
    const countRow = vData[r + 1];
    if (countRow) {
      for (let c = 1; c < vHeaders.length; c++) {
        const w = clean(vHeaders[c]);
        if (w && num(w) >= 12) vertCounts[w] = num(countRow[c]);
      }
    }
    break;
  }
}

// ── Parse Snow - Changers ──
const chSheet = findSheet("Snow - Changers");
const chData = sheetToArray(chSheet);

// Wind buckets
const windLoadBuckets = {};
for (let r = 0; r < Math.min(6, chData.length); r++) {
  const row = chData[r];
  if (!row) continue;
  if (r <= 3 && chData[r + 2]) {
    for (let c = 0; c < row.length; c++) {
      const inp = num(row[c]);
      const buck = num(chData[r + 2]?.[c]);
      if (inp >= 85 && inp <= 200 && buck >= 100 && buck <= 200) {
        windLoadBuckets[String(inp)] = buck;
      }
    }
  }
}

// Height classification
const heightClassification = {};
const feetUsedByHeight = {};
for (let r = 15; r < Math.min(35, chData.length); r++) {
  const row = chData[r];
  if (!row) continue;
  let smtCount = 0;
  for (let c = 0; c < row.length; c++) {
    const v = clean(row[c]);
    if (v === "S" || v === "M" || v === "T") smtCount++;
  }
  if (smtCount >= 5) {
    const heightRow = chData[r - 1];
    const feetRow = chData[r + 1];
    for (let c = 0; c < row.length; c++) {
      const prefix = clean(row[c]);
      if (prefix !== "S" && prefix !== "M" && prefix !== "T") continue;
      const h = num(heightRow?.[c]);
      if (h >= 1 && h <= 30) {
        heightClassification[String(h)] = prefix === "S" ? 0 : prefix === "M" ? 1 : 2;
        if (feetRow) feetUsedByHeight[String(h)] = num(feetRow[c]);
      }
    }
    break;
  }
}

// State pricing
const pieTrussPrice = {};
const trussPriceByWidthState = {};
const channelPriceByState = {};
const tubingPriceByState = {};
let stateCodeRow = -1;
let stateCodes = [];
for (let r = 50; r < Math.min(75, chData.length); r++) {
  const row = chData[r];
  if (!row) continue;
  let codeCount = 0;
  const codes = [];
  for (let c = 1; c < row.length; c++) {
    const v = clean(row[c]);
    if (v.match(/^[A-Z]{2}$/)) { codeCount++; codes.push(v); }
    else codes.push("");
  }
  if (codeCount >= 5) {
    stateCodeRow = r;
    stateCodes = codes;
    break;
  }
}

if (stateCodeRow >= 0) {
  for (let r = stateCodeRow + 1; r < Math.min(stateCodeRow + 12, chData.length); r++) {
    const row = chData[r];
    if (!row) continue;
    const label = clean(row[0]).toLowerCase();
    if (label.includes("wide") || label.includes("truss")) {
      const wm = label.match(/(\d+)/);
      if (wm) {
        const w = wm[1];
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
        const st = stateCodes[c - 1];
        if (!st) continue;
        const price = num(row[c]);
        if (price > 0 && price < 100) pieTrussPrice[st] = price;
      }
    } else if (label.includes("channel")) {
      for (let c = 1; c < row.length && c - 1 < stateCodes.length; c++) {
        const st = stateCodes[c - 1];
        if (!st) continue;
        const price = num(row[c]);
        if (price >= 1 && price <= 10) channelPriceByState[st] = price;
      }
    } else if (label.includes("tubing") || label.includes("tube")) {
      for (let c = 1; c < row.length && c - 1 < stateCodes.length; c++) {
        const st = stateCodes[c - 1];
        if (!st) continue;
        const price = num(row[c]);
        if (price >= 1 && price <= 10) tubingPriceByState[st] = price;
      }
    }
  }
}

// ══════════════════════════════════════════════════
// SIMULATE ENGINE: 24x30x10, 30GL, 90mph, AZ, AFV, Enclosed
// ══════════════════════════════════════════════════
console.log("=== ENGINE TRACE: 24x30x10, 30GL, 90mph, AZ, AFV, Enclosed ===\n");

const width = 24, length = 30, height = 10, state = "AZ";
const snowLoad = "30GL";
const roofKey = "AFV";
const sidesOrientation = "vertical";
const sidesCoverage = "fully_enclosed";
const sidesQty = 2;
const endsQty = 2;
const endsOrientation = "vertical";

// Step 1
const bw = windLoadBuckets["90"] || 105;
console.log("Wind buckets:", JSON.stringify(windLoadBuckets));
console.log("Bucketed wind:", bw);

const hc = heightClassification[String(height)];
const hp = hc === 0 ? "S" : hc === 1 ? "M" : "T";
console.log("Height class:", JSON.stringify(heightClassification));
console.log("Height prefix:", hp);

const sc = hp + "-" + snowLoad;
const isEnclosed = sidesCoverage === "fully_enclosed" && endsQty >= 2;
const ck = (isEnclosed ? "E" : "O") + "-" + bw + "-" + width + "-" + roofKey;
console.log("Snow code:", sc, "Config key:", ck);

// Step 2: Trusses
const ts = trussSpacing[sc]?.[ck] ?? 0;
console.log("\n--- TRUSSES ---");
console.log("Spacing lookup:", sc, "→", ck, "=", ts);
if (ts > 0) {
  const needed = Math.ceil((length * 12) / ts) + 1;
  const wsk = width + "-" + state;
  const orig = trussCounts[wsk]?.[String(length)] ?? 0;
  console.log("Needed:", needed, "Original key:", wsk + "/" + length, "=", orig);
  console.log("Extra:", Math.max(0, needed - orig));

  if (needed - orig > 0) {
    const trPrice = trussPriceByWidthState[state]?.[String(width)] ?? 190;
    const fu = feetUsedByHeight[String(height)] ?? 0;
    const pp = pieTrussPrice[state] ?? 15;
    const legSurcharge = fu * pp;
    console.log("Truss price:", trPrice, "FeetUsed:", fu, "PiePrice:", pp, "LegSurcharge:", legSurcharge);
    console.log("Truss cost:", (needed - orig) * (trPrice + legSurcharge));
  }
} else {
  console.log("No truss spacing found → 0 cost");
}

// Step 3: Hat Channels
const bts = 60;
const hck = bts + "-" + snowLoad;
const hcs = hcSpacing[hck]?.[String(bw)] ?? 0;
console.log("\n--- HAT CHANNELS ---");
console.log("HC key:", hck, "Wind:", bw, "→ spacing:", hcs);
if (hcs > 0) {
  const bar = (width + 2) / 2 * 12;
  const per = Math.ceil(bar / hcs) + 1;
  const total = per * 2;
  let origHC = hcCounts[state]?.[String(width)] ?? 0;
  if (origHC === 0) origHC = hcCounts[String(width)]?.[state] ?? 0;
  console.log("Bar:", bar, "PerSide:", per, "Total:", total, "Original:", origHC);
  console.log("Extra:", Math.max(0, total - origHC));
}

// Step 4: Girts
const gs = girtSpacing[String(bts)]?.[String(bw)] ?? 0;
console.log("\n--- GIRTS ---");
console.log("Gate: enclosed=" + isEnclosed, "sidesOrientation=" + sidesOrientation);
console.log("Girt spacing:", gs, "(key:", bts, "/", bw, ")");
console.log("Girt counts by height:", JSON.stringify(girtCountsByHeight));
if (gs > 0 && isEnclosed && sidesOrientation === "vertical") {
  const gr = Math.ceil((height * 12) / gs) + 1;
  let origG = girtCountsByHeight[String(height)] || 0;
  if (origG === 0) {
    for (const [k, v] of Object.entries(girtCountsByHeight)) {
      const p = k.split("-").map(Number);
      if (p.length === 2 && height >= p[0] && height <= p[1]) { origG = v; break; }
    }
  }
  if (origG === 0) origG = height <= 11 ? 3 : height <= 17 ? 4 : 5;
  console.log("Needed:", gr, "Original:", origG, "Extra:", Math.max(0, gr - origG));
}

// Step 5: Verticals
console.log("\n--- VERTICALS ---");
// The engine does: verticalSpacing[String(height)][String(wind)]
// But readMatrix without transpose stores as matrix[colHeader][rowKey]
// So if headers are heights (6,7,8,...20) and rows are wind speeds (105,...180),
// it would be vertSpacing["10"]["105"]
const vs1 = vertSpacing[String(height)]?.[String(bw)] ?? 0;
const vs2 = vertSpacing[String(bw)]?.[String(height)] ?? 0;
console.log("vertSpacing[height][wind]:", vs1);
console.log("vertSpacing[wind][height]:", vs2);
console.log("Vert spacing matrix sample keys:", Object.keys(vertSpacing).slice(0, 8));
const firstKey = Object.keys(vertSpacing)[0];
if (firstKey) console.log("First key sub-keys:", Object.keys(vertSpacing[firstKey]).slice(0, 5));

const vs = vs1 || vs2;
if (vs > 0) {
  const needed = Math.ceil((width * 12) / vs) + 1;
  const orig = vertCounts[String(width)] ?? 0;
  console.log("Needed:", needed, "Original:", orig, "Extra:", Math.max(0, needed - orig));
}

console.log("\nVert counts:", JSON.stringify(vertCounts));
console.log("\n=== EXPECTED: ALL EXTRA = 0 → $0 ENGINEERING ===");
