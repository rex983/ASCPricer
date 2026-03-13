/**
 * Trace script: 22x50x12, AFV roof, Horizontal sides, 30GL snow, 90mph wind, AZ
 *
 * Purpose: Identify why the app shows $280 when the Excel spreadsheet shows $0 ($-)
 * for this exact configuration.
 */

import * as XLSX from './node_modules/xlsx/xlsx.mjs';
import { readFileSync } from 'fs';

const FILE = 'C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx';

// ── Utility functions (matching the app) ──
function num(v) {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    const n = parseFloat(v.replace(/[$,]/g, ''));
    return isNaN(n) ? 0 : n;
  }
  return 0;
}

function cleanHeader(v) {
  return String(v ?? '').trim().replace(/\s+/g, ' ');
}

function nearestBucket(value, buckets) {
  let nearest = buckets[0];
  for (const bucket of buckets) {
    if (bucket <= value) nearest = bucket;
    else break;
  }
  return nearest;
}

function sheetToArray(ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
}

// ── Load workbook ──
const buf = readFileSync(FILE);
const wb = XLSX.read(buf, { type: 'buffer' });

console.log('=== SPREADSHEET SHEETS ===');
console.log(wb.SheetNames.join('\n'));
console.log();

// ── Configuration ──
const WIDTH = 22;
const LENGTH = 50;
const HEIGHT = 12;
const ROOF_KEY = 'AFV';
const SNOW_LOAD = '30GL';
const WIND_MPH = 90;
const STATE = 'AZ';
const SIDES_ORIENTATION = 'horizontal'; // Horizontal panels
const SIDES_COVERAGE = 'fully_enclosed';
const SIDES_QTY = 2;
const ENDS_QTY = 2;
const ENDS_ORIENTATION = 'vertical';

const WIND_LOAD_CATEGORIES = [105, 115, 130, 140, 155, 165, 180];
const TRUSS_SPACING_BUCKETS = [36, 42, 48, 54, 60];

console.log('=== INPUT CONFIGURATION ===');
console.log(`Width: ${WIDTH}, Length: ${LENGTH}, Height: ${HEIGHT}`);
console.log(`Roof: ${ROOF_KEY}, Snow: ${SNOW_LOAD}, Wind: ${WIND_MPH}mph`);
console.log(`State: ${STATE}, Sides: ${SIDES_ORIENTATION}, Coverage: ${SIDES_COVERAGE}`);
console.log();

// ══════════════════════════════════════════════════════════════
// STEP 1: Wind Bucketing (Snow - Changers, rows 0-1)
// ══════════════════════════════════════════════════════════════
console.log('=== STEP 1: WIND BUCKETING ===');
const changersSheet = wb.Sheets['Snow - Changers'];
if (!changersSheet) {
  console.error('ERROR: "Snow - Changers" sheet not found!');
  process.exit(1);
}
const changersData = sheetToArray(changersSheet);

// Dump rows 0-5 for wind bucketing analysis
console.log('Snow - Changers rows 0-5:');
for (let r = 0; r <= Math.min(5, changersData.length - 1); r++) {
  const row = changersData[r];
  const cells = row.slice(0, 25).map((v, i) => `[${i}]=${v}`).join(', ');
  console.log(`  Row ${r}: ${cells}`);
}

// Parse wind load buckets per the app logic
const windLoadBuckets = {};
if (changersData[0] && changersData[1]) {
  for (let c = 1; c < changersData[0].length; c++) {
    const inputMph = num(changersData[0][c]);
    const bucketMph = num(changersData[1][c]);
    if (inputMph >= 80 && inputMph <= 200 && bucketMph >= 100 && bucketMph <= 200) {
      windLoadBuckets[String(inputMph)] = bucketMph;
    }
  }
}

console.log('\nParsed wind load buckets (relevant):');
for (const mph of [85, 90, 95, 100, 105, 110, 115, 120, 130]) {
  console.log(`  ${mph}mph → ${windLoadBuckets[String(mph)] ?? 'NOT FOUND'}`);
}

// What does 90mph bucket to?
const bucketedFromLookup = windLoadBuckets[String(WIND_MPH)];
const bucketedFromNearest = nearestBucket(WIND_MPH, WIND_LOAD_CATEGORIES);

// App logic: bucketWind checks windLoadBuckets first, falls back to nearestBucket
let bucketedWind;
if (bucketedFromLookup && bucketedFromLookup > 0) {
  bucketedWind = bucketedFromLookup;
  console.log(`\n90mph → windLoadBuckets["90"] = ${bucketedFromLookup} (from lookup)`);
} else {
  bucketedWind = bucketedFromNearest;
  console.log(`\n90mph → NOT in windLoadBuckets, nearestBucket(90, [105,115,...]) = ${bucketedFromNearest}`);
  console.log('  NOTE: nearestBucket rounds DOWN. 90 < 105, so no bucket <= 90 exists!');
  console.log(`  Actually nearestBucket returns first bucket (${WIND_LOAD_CATEGORIES[0]}) as fallback`);
}

console.log(`\nFINAL bucketed wind: ${bucketedWind}`);
console.log();

// ══════════════════════════════════════════════════════════════
// STEP 2: Height Classification (Snow - Changers, rows ~25-26)
// ══════════════════════════════════════════════════════════════
console.log('=== STEP 2: HEIGHT CLASSIFICATION ===');

const heightClassification = {};
const feetUsedByHeight = {};

// Find the S/M/T row
let smtRowIdx = -1;
for (let r = 15; r < Math.min(35, changersData.length); r++) {
  const row = changersData[r];
  if (!row) continue;
  let smtCount = 0;
  for (let c = 0; c < row.length; c++) {
    const v = cleanHeader(row[c]);
    if (v === 'S' || v === 'M' || v === 'T') smtCount++;
  }
  if (smtCount >= 5) {
    smtRowIdx = r;
    const heightRow = changersData[r - 1];
    const feetRow = changersData[r + 1];

    console.log(`Found S/M/T at row ${r}:`);
    console.log(`  Height row (${r-1}): ${heightRow?.slice(0, 25).join(', ')}`);
    console.log(`  S/M/T row  (${r}):   ${row.slice(0, 25).join(', ')}`);
    console.log(`  FeetUsed   (${r+1}): ${feetRow?.slice(0, 25).join(', ')}`);

    if (heightRow) {
      for (let c = 0; c < row.length; c++) {
        const prefix = cleanHeader(row[c]);
        if (prefix !== 'S' && prefix !== 'M' && prefix !== 'T') continue;
        const h = num(heightRow[c]);
        if (h >= 1 && h <= 30) {
          heightClassification[String(h)] = prefix === 'S' ? 0 : prefix === 'M' ? 1 : 2;
          if (feetRow) {
            feetUsedByHeight[String(h)] = num(feetRow[c]);
          }
        }
      }
    }
    break;
  }
}

const heightVal = heightClassification[String(HEIGHT)];
const heightPrefix = heightVal === 0 ? 'S' : heightVal === 1 ? 'M' : heightVal === 2 ? 'T' : '??';
console.log(`\nHeight ${HEIGHT} → classification value: ${heightVal}, prefix: "${heightPrefix}"`);
console.log(`FeetUsed for height ${HEIGHT}: ${feetUsedByHeight[String(HEIGHT)]}`);
console.log();

// ══════════════════════════════════════════════════════════════
// STEP 3: Snow Code
// ══════════════════════════════════════════════════════════════
const snowCode = `${heightPrefix}-${SNOW_LOAD}`;
console.log(`=== STEP 3: SNOW CODE ===`);
console.log(`Snow code: "${heightPrefix}-${SNOW_LOAD}" = "${snowCode}"`);
console.log();

// ══════════════════════════════════════════════════════════════
// STEP 4: E/O determination
// ══════════════════════════════════════════════════════════════
console.log('=== STEP 4: ENCLOSED/OPEN ===');
const isEnclosed = SIDES_COVERAGE === 'fully_enclosed' && SIDES_ORIENTATION === 'vertical' && SIDES_QTY >= 2;
console.log(`Sides: ${SIDES_ORIENTATION}, Coverage: ${SIDES_COVERAGE}, SidesQty: ${SIDES_QTY}`);
console.log(`isEnclosed = (fully_enclosed && vertical && qty>=2) = ${isEnclosed}`);
console.log(`Horizontal panels → treated as OPEN (O) for engineering`);
const enclosure = isEnclosed ? 'E' : 'O';
console.log(`Enclosure: "${enclosure}"`);
console.log();

// ══════════════════════════════════════════════════════════════
// STEP 5: Config Key
// ══════════════════════════════════════════════════════════════
const configKey = `${enclosure}-${bucketedWind}-${WIDTH}-${ROOF_KEY}`;
console.log('=== STEP 5: CONFIG KEY ===');
console.log(`Config key: "${enclosure}-${bucketedWind}-${WIDTH}-${ROOF_KEY}" = "${configKey}"`);
console.log();

// ══════════════════════════════════════════════════════════════
// STEP 6: Truss Spacing Lookup (Snow - Truss Spacing)
// ══════════════════════════════════════════════════════════════
console.log('=== STEP 6: TRUSS SPACING ===');
const trussSpacingSheet = wb.Sheets['Snow - Truss Spacing'];
if (!trussSpacingSheet) {
  console.error('ERROR: "Snow - Truss Spacing" sheet not found!');
  process.exit(1);
}
const tsData = sheetToArray(trussSpacingSheet);

// Dump first few rows and headers
console.log('Snow - Truss Spacing headers (first 20):');
const tsHeaders = tsData[0]?.slice(0, 20) || [];
console.log(`  ${tsHeaders.join(' | ')}`);

// The app reads this transposed: matrix[snowCode][configKey]
// readMatrix with transpose=true means matrix[rowKey][colKey]
// Row keys are in col 0 (snow codes like "S-30GL"), col headers are config keys
const trussSpacingMatrix = {};
const tsHeaderRow = tsData[0] || [];
for (let r = 1; r < tsData.length; r++) {
  const row = tsData[r];
  if (!row) break;
  const rowKey = cleanHeader(row[0]);
  if (!rowKey || rowKey === '0') continue;
  trussSpacingMatrix[rowKey] = {};
  for (let c = 1; c < row.length; c++) {
    const colKey = cleanHeader(tsHeaderRow[c]);
    if (!colKey || colKey === '0') continue;
    trussSpacingMatrix[rowKey][colKey] = num(row[c]);
  }
}

// Look for our snow code row
console.log(`\nLooking for snow code row "${snowCode}":`);
if (trussSpacingMatrix[snowCode]) {
  console.log(`  Found! Keys available: ${Object.keys(trussSpacingMatrix[snowCode]).slice(0, 15).join(', ')}`);
} else {
  console.log(`  NOT FOUND! Available row keys:`);
  const rowKeys = Object.keys(trussSpacingMatrix);
  console.log(`  ${rowKeys.join(', ')}`);
}

// Look for config key column
console.log(`\nLooking for config key column "${configKey}":`);
const allColKeys = new Set();
for (const rk of Object.keys(trussSpacingMatrix)) {
  for (const ck of Object.keys(trussSpacingMatrix[rk])) {
    allColKeys.add(ck);
  }
}
const colKeysArr = [...allColKeys].sort();
const matchingCols = colKeysArr.filter(k => k.includes('22') || k.includes('O-'));
console.log(`  Columns containing "22" or starting with "O-":`);
console.log(`  ${matchingCols.join(', ')}`);

// Actual lookup
const trussSpacing = trussSpacingMatrix[snowCode]?.[configKey] ?? 0;
console.log(`\ntrussSpacing[${snowCode}][${configKey}] = ${trussSpacing}`);

if (trussSpacing === 0) {
  console.log('*** TRUSS SPACING = 0 → This means "Contact Engineering" (return -1) ***');
  console.log('*** BUT the app checks: if (trussSpacing === 0 || trussSpacing < 18) return -1 ***');
  console.log('*** Wait - the app returns -1 for "Contact Engineering", not $280. ***');
  console.log('*** Let us check what actually happens... ***');
}

// Also check nearby configs to see the pattern
console.log('\n--- Nearby config lookups for snow code "' + snowCode + '" ---');
for (const wind of [105, 115, 130]) {
  for (const enc of ['O', 'E']) {
    const key = `${enc}-${wind}-${WIDTH}-${ROOF_KEY}`;
    const val = trussSpacingMatrix[snowCode]?.[key] ?? 'MISSING';
    console.log(`  ${key} → ${val}`);
  }
}

// Check what the ACTUAL wind bucket for 90 gives us
console.log('\n--- What about if 90mph does NOT bucket to any standard category? ---');
console.log(`  nearestBucket(90, [105,115,130,140,155,165,180]):`);
console.log(`  Walking buckets: 105 > 90, so we never enter the loop body.`);
console.log(`  Returns buckets[0] = 105`);
console.log(`  So 90mph → 105 in the app!`);

// The actual config key used by the app
const appConfigKey = `O-105-${WIDTH}-${ROOF_KEY}`;
const appTrussSpacing = trussSpacingMatrix[snowCode]?.[appConfigKey] ?? 0;
console.log(`\nApp actually uses: trussSpacing[${snowCode}][${appConfigKey}] = ${appTrussSpacing}`);
console.log();

// ══════════════════════════════════════════════════════════════
// STEP 7: Original Truss Count (Snow - Trusses)
// ══════════════════════════════════════════════════════════════
console.log('=== STEP 7: ORIGINAL TRUSS COUNT ===');
// Sheet name has trailing space in the spreadsheet
const trussCountSheet = wb.Sheets['Snow - Trusses'] || wb.Sheets['Snow - Trusses '];
if (!trussCountSheet) {
  console.error('ERROR: "Snow - Trusses" sheet not found!');
  process.exit(1);
}
const tcData = sheetToArray(trussCountSheet);

// readTrussCounts uses readMatrix WITHOUT transpose → matrix[colKey][rowKey]
// So: trussCounts["{width}-{state}"]["{length}"]
const trussCounts = {};
const tcHeaders = tcData[0] || [];
for (let r = 1; r < tcData.length; r++) {
  const row = tcData[r];
  if (!row) break;
  const rowKey = cleanHeader(row[0]);
  if (!rowKey || rowKey === '0') continue;
  for (let c = 1; c < row.length; c++) {
    const colKey = cleanHeader(tcHeaders[c]);
    if (!colKey || colKey === '0') continue;
    if (!trussCounts[colKey]) trussCounts[colKey] = {};
    trussCounts[colKey][rowKey] = num(row[c]);
  }
}

// Resolve state
let resolvedState = STATE;
const widthStateKey = `${WIDTH}-${STATE}`;
if (!trussCounts[widthStateKey]) {
  const prefix = `${WIDTH}-`;
  for (const key of Object.keys(trussCounts)) {
    if (key.startsWith(prefix)) {
      resolvedState = key.slice(prefix.length);
      break;
    }
  }
}
const finalWidthStateKey = `${WIDTH}-${resolvedState}`;

console.log(`Looking for "${widthStateKey}" in truss counts...`);
console.log(`Available keys with width ${WIDTH}: ${Object.keys(trussCounts).filter(k => k.startsWith(WIDTH + '-')).join(', ')}`);
console.log(`Resolved state: ${resolvedState}, using key: "${finalWidthStateKey}"`);

const originalTrusses = trussCounts[finalWidthStateKey]?.['50'] ?? 0;
console.log(`Original trusses for ${finalWidthStateKey} at length 50: ${originalTrusses}`);

// Calculate needed trusses (if truss spacing > 0)
const useTrussSpacing = appTrussSpacing; // what the app actually uses
if (useTrussSpacing > 0) {
  const lengthInches = LENGTH * 12;
  const trussesNeeded = Math.ceil(lengthInches / useTrussSpacing) + 1;
  const extraTrusses = Math.max(0, trussesNeeded - originalTrusses);
  console.log(`Trusses needed: ceil(${lengthInches} / ${useTrussSpacing}) + 1 = ${trussesNeeded}`);
  console.log(`Extra trusses: max(0, ${trussesNeeded} - ${originalTrusses}) = ${extraTrusses}`);
}
console.log();

// ══════════════════════════════════════════════════════════════
// STEP 8: Hat Channel Spacing (Snow - Hat Channels)
// ══════════════════════════════════════════════════════════════
console.log('=== STEP 8: HAT CHANNEL SPACING ===');
const hcSheet = wb.Sheets['Snow - Hat Channels'];
if (!hcSheet) {
  console.error('ERROR: "Snow - Hat Channels" sheet not found!');
  process.exit(1);
}
const hcData = sheetToArray(hcSheet);

// Parse left side (spacing)
const hcHeaders = hcData[0] || [];
console.log('Hat Channel headers:', hcHeaders.slice(0, 15).join(' | '));

const hcSpacing = {};
// Find right section boundary
let hcRightStart = -1;
for (let c = 8; c < hcHeaders.length; c++) {
  const h = cleanHeader(hcHeaders[c]);
  if (h.toLowerCase().includes('original')) {
    hcRightStart = c;
    break;
  }
}
if (hcRightStart < 0) {
  for (let c = 8; c < hcHeaders.length; c++) {
    const h = cleanHeader(hcHeaders[c]);
    if (['105','115','130','140','155','165','180','','0'].includes(h)) continue;
    let stateCount = 0;
    for (let r = 1; r < Math.min(10, hcData.length); r++) {
      const v = cleanHeader(hcData[r]?.[c] ?? '');
      if (v.match(/^[A-Z]{2}$/)) stateCount++;
    }
    if (stateCount >= 3) {
      hcRightStart = c;
      break;
    }
  }
}

const spacingEndCol = hcRightStart > 0 ? hcRightStart : 8;
for (let r = 1; r < hcData.length; r++) {
  const row = hcData[r];
  if (!row) break;
  const rowKey = cleanHeader(row[0]);
  if (!rowKey) break;
  if (!hcSpacing[rowKey]) hcSpacing[rowKey] = {};
  for (let c = 1; c < spacingEndCol; c++) {
    const colKey = cleanHeader(hcHeaders[c]);
    if (!colKey || colKey === '0') continue;
    hcSpacing[rowKey][colKey] = num(row[c]);
  }
}

// Parse right side (original counts)
const hcOriginalCounts = {};
if (hcRightStart > 0) {
  let stateCol = hcRightStart;
  let widthStartCol = hcRightStart + 1;
  let widthHeaders = [];
  for (let c = widthStartCol; c < hcHeaders.length; c++) {
    const h = cleanHeader(hcHeaders[c]);
    if (h && num(h) >= 12 && num(h) <= 30) {
      widthHeaders.push(h);
    } else {
      widthHeaders.push('');
    }
  }

  for (let r = 1; r < hcData.length; r++) {
    const row = hcData[r];
    if (!row) break;
    const state = cleanHeader(row[stateCol]);
    if (!state || !state.match(/^[A-Z]{2}$/)) continue;
    if (!hcOriginalCounts[state]) hcOriginalCounts[state] = {};
    for (let c = widthStartCol; c < row.length; c++) {
      const widthKey = widthHeaders[c - widthStartCol];
      if (!widthKey) continue;
      const count = num(row[c]);
      if (count > 0) hcOriginalCounts[state][widthKey] = count;
    }
  }
}

// HC lookup: row key is "{bucketedTrussSpacing}-{snowLoad}" — NO height prefix
const actualTrussSpacing = useTrussSpacing > 0 ? useTrussSpacing : 60;
const bucketedTrussSpacing = nearestBucket(actualTrussSpacing, TRUSS_SPACING_BUCKETS);
const hcRowKey = `${bucketedTrussSpacing}-${SNOW_LOAD}`;

console.log(`\nTruss spacing for HC lookup: ${actualTrussSpacing} → bucketed: ${bucketedTrussSpacing}`);
console.log(`HC row key: "${hcRowKey}" (NO height prefix per app logic)`);
console.log(`HC col key (wind): "${bucketedWind}"`);

// Also try with height prefix to see difference
const hcRowKeyWithPrefix = `${bucketedTrussSpacing}-${snowCode}`;
console.log(`HC row key WITH height prefix would be: "${hcRowKeyWithPrefix}"`);

console.log(`\nAvailable HC row keys (first 20): ${Object.keys(hcSpacing).slice(0, 20).join(', ')}`);

const hatChannelSpacing = hcSpacing[hcRowKey]?.[String(bucketedWind)] ?? 0;
const hatChannelSpacingWithPrefix = hcSpacing[hcRowKeyWithPrefix]?.[String(bucketedWind)] ?? 0;
console.log(`hatChannelSpacing["${hcRowKey}"]["${bucketedWind}"] = ${hatChannelSpacing}`);
console.log(`hatChannelSpacing["${hcRowKeyWithPrefix}"]["${bucketedWind}"] = ${hatChannelSpacingWithPrefix}`);

// ══════════════════════════════════════════════════════════════
// STEP 9: Original HC Count
// ══════════════════════════════════════════════════════════════
console.log('\n=== STEP 9: ORIGINAL HAT CHANNEL COUNT ===');
let originalChannels = hcOriginalCounts[resolvedState]?.[String(WIDTH)] ?? 0;
if (originalChannels === 0) {
  // Try width as row key
  originalChannels = hcOriginalCounts[String(WIDTH)]?.[resolvedState] ?? 0;
}
console.log(`Original HC for state=${resolvedState}, width=${WIDTH}: ${originalChannels}`);
console.log(`HC original counts available:`, JSON.stringify(hcOriginalCounts, null, 2));
console.log();

// ══════════════════════════════════════════════════════════════
// STEP 10: Girt Spacing (Snow - Girts)
// ══════════════════════════════════════════════════════════════
console.log('=== STEP 10: GIRT SPACING ===');
const girtSheet = wb.Sheets['Snow - Girts'] || wb.Sheets['Snow - Girts '];
if (!girtSheet) {
  console.error('ERROR: "Snow - Girts" sheet not found!');
  process.exit(1);
}
const girtData = sheetToArray(girtSheet);

// Parse girt spacing
const girtHeaders = girtData[0] || [];
const girtSpacing = {};
const girtCountsByHeight = {};

let girtRightStart = -1;
for (let c = 1; c < girtHeaders.length; c++) {
  const h = cleanHeader(girtHeaders[c]);
  if (h === '' || h === '0') {
    const next = cleanHeader(girtHeaders[c + 1]);
    if (next && !['105','115','130','140','155','165','180'].includes(next)) {
      girtRightStart = c + 1;
      break;
    }
  }
}

const girtSpacingEndCol = girtRightStart > 0 ? girtRightStart : girtHeaders.length;
for (let r = 1; r < girtData.length; r++) {
  const row = girtData[r];
  if (!row) break;
  const rowKey = cleanHeader(row[0]);
  if (!rowKey) break;
  if (!girtSpacing[rowKey]) girtSpacing[rowKey] = {};
  for (let c = 1; c < girtSpacingEndCol; c++) {
    const colKey = cleanHeader(girtHeaders[c]);
    if (!colKey || colKey === '0') continue;
    girtSpacing[rowKey][colKey] = num(row[c]);
  }
}

if (girtRightStart > 0) {
  for (let r = 1; r < girtData.length; r++) {
    const row = girtData[r];
    if (!row) break;
    const heightKey = cleanHeader(row[girtRightStart]);
    const count = num(row[girtRightStart + 1]);
    if (!heightKey) continue;
    if (count > 0) girtCountsByHeight[heightKey] = count;
  }
}

console.log(`Girt spacing matrix keys: ${Object.keys(girtSpacing).join(', ')}`);
console.log(`Girt counts by height: ${JSON.stringify(girtCountsByHeight)}`);

// App: girts only if isEnclosed AND vertical panels
const girtsNeeded = isEnclosed && SIDES_ORIENTATION === 'vertical';
console.log(`\nGirts needed? isEnclosed=${isEnclosed} && vertical=${SIDES_ORIENTATION === 'vertical'} → ${girtsNeeded}`);
console.log('(Horizontal panels → girts are SKIPPED)');

// But let's check the spacing anyway for reference
const girtLookupKey = String(bucketedTrussSpacing);
const girtSpacingVal = girtSpacing[girtLookupKey]?.[String(bucketedWind)] ?? 0;
console.log(`Girt spacing[${girtLookupKey}][${bucketedWind}] = ${girtSpacingVal} (not used since girtsNeeded=false)`);
console.log();

// ══════════════════════════════════════════════════════════════
// STEP 11: Original Girt Count
// ══════════════════════════════════════════════════════════════
console.log('=== STEP 11: ORIGINAL GIRT COUNT ===');
// resolveOriginalGirts logic
let originalGirts = girtCountsByHeight[String(HEIGHT)] ?? 0;
if (originalGirts === 0) {
  for (const [key, count] of Object.entries(girtCountsByHeight)) {
    const parts = key.split('-').map(Number);
    if (parts.length === 2 && HEIGHT >= parts[0] && HEIGHT <= parts[1]) {
      originalGirts = count;
      break;
    }
  }
}
if (originalGirts === 0) {
  originalGirts = HEIGHT <= 11 ? 3 : HEIGHT <= 17 ? 4 : 5;
}
console.log(`Original girts for height ${HEIGHT}: ${originalGirts}`);
console.log('(Not used since horizontal panels → no girt calculation)');
console.log();

// ══════════════════════════════════════════════════════════════
// STEP 12: Vertical Spacing (Snow - Verticals)
// ══════════════════════════════════════════════════════════════
console.log('=== STEP 12: VERTICAL SPACING ===');
const vertSheet = wb.Sheets['Snow - Verticals'];
if (!vertSheet) {
  console.error('ERROR: "Snow - Verticals" sheet not found!');
  process.exit(1);
}
const vertData = sheetToArray(vertSheet);
const vertHeaders = vertData[0] || [];

// readVerticals uses readMatrix WITHOUT transpose: matrix[colKey][rowKey]
// So: verticalSpacing[colHeader][rowKey] where row keys come from col 0
// Columns are height values, rows are wind speeds? Or vice versa?
// Let's dump the data to see
console.log('Verticals sheet headers:', vertHeaders.slice(0, 15).join(' | '));
console.log('First 10 rows:');
for (let r = 0; r < Math.min(15, vertData.length); r++) {
  console.log(`  Row ${r}: ${(vertData[r] || []).slice(0, 15).join(' | ')}`);
}

// readMatrix WITHOUT transpose: matrix[colHeader][rowKey]
const vertSpacing = {};
for (let r = 1; r < vertData.length; r++) {
  const row = vertData[r];
  if (!row) break;
  const rowKey = cleanHeader(row[0]);
  if (!rowKey || rowKey === '0') continue;
  for (let c = 1; c < row.length; c++) {
    const colKey = cleanHeader(vertHeaders[c]);
    if (!colKey || colKey === '0') continue;
    if (!vertSpacing[colKey]) vertSpacing[colKey] = {};
    vertSpacing[colKey][rowKey] = num(row[c]);
  }
}

// The app uses: lookupMatrix(verticalSpacing, String(height), String(bucketedWind))
// With non-transposed matrix: matrix[colKey][rowKey] → vertSpacing[String(height)][String(bucketedWind)]
// Wait, the app passes (matrix, String(height), String(bucketedWind))
// lookupMatrix returns matrix[rowKey]?.[colKey] = matrix[String(height)]?.[String(bucketedWind)]
// But the matrix from readMatrix(non-transposed) is matrix[colKey][rowKey]
// So this lookup is: vertSpacing[String(HEIGHT)]?.[String(bucketedWind)]

console.log(`\nVertical spacing lookup: vertSpacing["${HEIGHT}"]["${bucketedWind}"]`);
console.log(`Available top-level keys: ${Object.keys(vertSpacing).join(', ')}`);

const verticalSpacing = vertSpacing[String(HEIGHT)]?.[String(bucketedWind)] ?? 0;
console.log(`Vertical spacing = ${verticalSpacing}`);

// Original vertical counts
const verticalCounts = {};
for (let r = 7; r < Math.min(20, vertData.length); r++) {
  const row = vertData[r];
  if (!row) continue;
  const label = cleanHeader(row[0]);
  if (label.toLowerCase().includes('original')) {
    const widthRow = row;
    const countRow = vertData[r + 1];
    console.log(`\nOriginal verticals label at row ${r}: "${label}"`);
    console.log(`  Width row: ${widthRow.slice(0, 15).join(', ')}`);
    console.log(`  Count row: ${(countRow || []).slice(0, 15).join(', ')}`);
    if (countRow) {
      for (let c = 1; c < widthRow.length; c++) {
        const w = cleanHeader(widthRow[c]);
        if (w && num(w) >= 12) {
          verticalCounts[w] = num(countRow[c]);
        }
      }
    }
    break;
  }
}

const originalVerticals = verticalCounts[String(WIDTH)] ?? 0;
console.log(`\nOriginal verticals for width ${WIDTH}: ${originalVerticals}`);

if (verticalSpacing > 0) {
  const widthInches = WIDTH * 12;
  const verticalsNeeded = Math.ceil(widthInches / verticalSpacing) + 1;
  const extraVerticals = Math.max(0, verticalsNeeded - originalVerticals);
  console.log(`Verticals needed: ceil(${widthInches} / ${verticalSpacing}) + 1 = ${verticalsNeeded}`);
  console.log(`Extra verticals: max(0, ${verticalsNeeded} - ${originalVerticals}) = ${extraVerticals}`);
}
console.log();

// ══════════════════════════════════════════════════════════════
// STEP 13: FULL CALCULATION (mimicking app logic exactly)
// ══════════════════════════════════════════════════════════════
console.log('=== STEP 13: FULL CALCULATION ===');
console.log('Reproducing calculateStandardSnowEngineering() step by step...\n');

// Recalculate with the EXACT app logic
const appBucketedWind = (() => {
  const b = windLoadBuckets[String(WIND_MPH)];
  if (b && b > 0) return b;
  return nearestBucket(WIND_MPH, WIND_LOAD_CATEGORIES);
})();

const appHeightPrefix = (() => {
  const val = heightClassification[String(HEIGHT)];
  if (val === 0) return 'S';
  if (val === 1) return 'M';
  if (val === 2) return 'T';
  if (HEIGHT <= 6) return 'S';
  if (HEIGHT <= 9) return 'M';
  return 'T';
})();

const appSnowCode = `${appHeightPrefix}-${SNOW_LOAD}`;
const appIsEnclosed = SIDES_COVERAGE === 'fully_enclosed' && SIDES_ORIENTATION === 'vertical' && SIDES_QTY >= 2;
const appEnclosure = appIsEnclosed ? 'E' : 'O';
const appFinalConfigKey = `${appEnclosure}-${appBucketedWind}-${WIDTH}-${ROOF_KEY}`;

console.log(`  bucketedWind: ${appBucketedWind}`);
console.log(`  heightPrefix: ${appHeightPrefix}`);
console.log(`  snowCode: ${appSnowCode}`);
console.log(`  isEnclosed: ${appIsEnclosed}`);
console.log(`  configKey: ${appFinalConfigKey}`);

// Step 2: Extra Trusses
const finalTrussSpacing = trussSpacingMatrix[appSnowCode]?.[appFinalConfigKey] ?? 0;
console.log(`\n  [TRUSSES] trussSpacing[${appSnowCode}][${appFinalConfigKey}] = ${finalTrussSpacing}`);

if (finalTrussSpacing === 0 || finalTrussSpacing < 18) {
  console.log(`  *** trussSpacing is ${finalTrussSpacing} → app returns -1 (Contact Engineering) ***`);
  console.log(`  *** But wait: the UI might display this differently... ***`);
}

let totalCost = 0;
let trussCost = 0;
let hcCost = 0;
let girtCost = 0;
let vertCost = 0;

if (finalTrussSpacing > 0 && finalTrussSpacing >= 18) {
  const lengthInches = LENGTH * 12;
  const trussesNeeded = Math.ceil(lengthInches / finalTrussSpacing) + 1;
  const origTr = trussCounts[finalWidthStateKey]?.['50'] ?? 0;
  const extraTrusses = Math.max(0, trussesNeeded - origTr);

  console.log(`  Trusses needed: ${trussesNeeded}, original: ${origTr}, extra: ${extraTrusses}`);

  if (extraTrusses > 0) {
    // Get truss price
    // trussPriceByWidthState parsed from Snow - Changers
    const pieTrussPrice = {};
    const trussPriceByWidthState = {};
    const channelPriceByState = {};
    const tubingPriceByState = {};

    // Re-parse per-state pricing section from changers
    let stateCodeRow = -1;
    let stateCodes = [];
    for (let r = 50; r < Math.min(75, changersData.length); r++) {
      const row = changersData[r];
      if (!row) continue;
      let codeCount = 0;
      const codes = [];
      for (let c = 1; c < row.length; c++) {
        const v = cleanHeader(row[c]);
        if (v.match(/^[A-Z]{2}$/)) {
          codeCount++;
          codes.push(v);
        } else {
          codes.push('');
        }
      }
      if (codeCount >= 5) {
        stateCodeRow = r;
        stateCodes = codes;
        break;
      }
    }

    if (stateCodeRow >= 0) {
      for (let r = stateCodeRow + 1; r < Math.min(stateCodeRow + 12, changersData.length); r++) {
        const row = changersData[r];
        if (!row) continue;
        const label = cleanHeader(row[0]).toLowerCase();

        if (label.includes('pie')) {
          for (let c = 1; c < row.length && c - 1 < stateCodes.length; c++) {
            const st = stateCodes[c - 1];
            if (!st) continue;
            pieTrussPrice[st] = num(row[c]);
          }
        } else if (label.includes('channel')) {
          for (let c = 1; c < row.length && c - 1 < stateCodes.length; c++) {
            const st = stateCodes[c - 1];
            if (!st) continue;
            channelPriceByState[st] = num(row[c]);
          }
        } else if (label.includes('tubing') || label.includes('tube')) {
          for (let c = 1; c < row.length && c - 1 < stateCodes.length; c++) {
            const st = stateCodes[c - 1];
            if (!st) continue;
            tubingPriceByState[st] = num(row[c]);
          }
        } else if (label.includes('wide') || label.includes('truss')) {
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
        }
      }
    }

    // getTrussPrice
    const stateRow = trussPriceByWidthState[resolvedState];
    let baseTrussPrice = 190;
    if (stateRow) {
      if (stateRow[String(WIDTH)] !== undefined) baseTrussPrice = stateRow[String(WIDTH)];
      else {
        for (const [rangeKey, price] of Object.entries(stateRow)) {
          const parts = rangeKey.split('-').map(Number);
          if (parts.length === 2 && WIDTH >= parts[0] && WIDTH <= parts[1]) {
            baseTrussPrice = price;
            break;
          }
        }
      }
    }

    const baseFeetUsed = feetUsedByHeight[String(HEIGHT)] ?? 0;
    const effectiveFeetUsed = (WIDTH >= 26 && HEIGHT < 13) ? baseFeetUsed * 2 : baseFeetUsed;
    const piePricePerFt = pieTrussPrice[resolvedState] || 15;
    const legSurcharge = effectiveFeetUsed * piePricePerFt;
    trussCost = extraTrusses * (baseTrussPrice + legSurcharge);

    console.log(`  baseTrussPrice: ${baseTrussPrice}, feetUsed: ${baseFeetUsed}, effectiveFeetUsed: ${effectiveFeetUsed}`);
    console.log(`  piePricePerFt: ${piePricePerFt}, legSurcharge: ${legSurcharge}`);
    console.log(`  TRUSS COST: ${extraTrusses} × (${baseTrussPrice} + ${legSurcharge}) = $${trussCost}`);
  }

  // Step 3: HC
  const bucketTS = nearestBucket(finalTrussSpacing, TRUSS_SPACING_BUCKETS);
  const hcKey = `${bucketTS}-${SNOW_LOAD}`;
  const hcSpacingVal = hcSpacing[hcKey]?.[String(appBucketedWind)] ?? 0;
  console.log(`\n  [HAT CHANNELS] hcSpacing[${hcKey}][${appBucketedWind}] = ${hcSpacingVal}`);

  if (hcSpacingVal > 0) {
    const barSize = (WIDTH + 2) / 2;
    const barInches = barSize * 12;
    const channelsPerSide = Math.ceil(barInches / hcSpacingVal) + 1;
    const totalChannelCalc = channelsPerSide * 2;

    let origHC = hcOriginalCounts[resolvedState]?.[String(WIDTH)] ?? 0;
    if (origHC === 0) origHC = hcOriginalCounts[String(WIDTH)]?.[resolvedState] ?? 0;

    const extraChannels = Math.max(0, totalChannelCalc - origHC);
    const channelPricePerFt = 2; // fallback
    const channelLength = LENGTH + 1;
    hcCost = extraChannels * channelPricePerFt * channelLength;

    console.log(`  barSize: ${barSize}, barInches: ${barInches}`);
    console.log(`  channelsPerSide: ${channelsPerSide}, total: ${totalChannelCalc}`);
    console.log(`  original: ${origHC}, extra: ${extraChannels}`);
    console.log(`  HC COST: ${extraChannels} × $${channelPricePerFt} × ${channelLength} = $${hcCost}`);
  }

  // Step 4: Girts (skipped for horizontal)
  console.log(`\n  [GIRTS] Skipped (horizontal panels → not enclosed for engineering)`);

  // Step 5: Verticals
  const vSpacing = vertSpacing[String(HEIGHT)]?.[String(appBucketedWind)] ?? 0;
  console.log(`\n  [VERTICALS] verticalSpacing[${HEIGHT}][${appBucketedWind}] = ${vSpacing}`);

  if (vSpacing > 0) {
    const widthInches = WIDTH * 12;
    const verticalsNeeded = Math.ceil(widthInches / vSpacing) + 1;
    const origVert = verticalCounts[String(WIDTH)] ?? 0;
    const extraVert = Math.max(0, verticalsNeeded - origVert);

    if (extraVert > 0 && ENDS_QTY > 0) {
      const tubPrice = 3; // fallback
      const ROOF_PITCH = 3/12;
      const roofRise = (ROOF_KEY === 'AFV' || ROOF_KEY === 'AFH') ? Math.ceil((WIDTH/2) * ROOF_PITCH) : 0;
      const peakHeight = HEIGHT + roofRise;
      const baseVertCost = extraVert * ENDS_QTY * tubPrice * peakHeight;
      const heightMult = HEIGHT >= 19 ? 3.0 : HEIGHT >= 16 ? 2.5 : HEIGHT >= 13 ? 2.0 : 1.0;
      vertCost = baseVertCost * heightMult;

      console.log(`  verticalsNeeded: ${verticalsNeeded}, original: ${origVert}, extra: ${extraVert}`);
      console.log(`  roofRise: ${roofRise}, peakHeight: ${peakHeight}`);
      console.log(`  baseVertCost: ${extraVert} × ${ENDS_QTY} × $${tubPrice} × ${peakHeight} = $${baseVertCost}`);
      console.log(`  heightMult: ${heightMult}`);
      console.log(`  VERT COST: $${baseVertCost} × ${heightMult} = $${vertCost}`);
    } else {
      console.log(`  verticalsNeeded: ${verticalsNeeded}, original: ${origVert}, extra: ${Math.max(0, verticalsNeeded - origVert)}`);
      console.log(`  No extra verticals needed or no ends`);
    }
  }

  totalCost = Math.round(trussCost + hcCost + girtCost + vertCost);
} else {
  console.log(`\n  Truss spacing is ${finalTrussSpacing} → Contact Engineering`);
  totalCost = -1;
}

console.log('\n' + '='.repeat(60));
console.log('=== FINAL RESULT ===');
console.log(`  Truss cost:   $${trussCost}`);
console.log(`  HC cost:      $${hcCost}`);
console.log(`  Girt cost:    $${girtCost}`);
console.log(`  Vert cost:    $${vertCost}`);
console.log(`  TOTAL:        $${totalCost}`);
console.log('='.repeat(60));

// ══════════════════════════════════════════════════════════════
// DIAGNOSTIC: Check spreadsheet's own math sheet for this config
// ══════════════════════════════════════════════════════════════
console.log('\n=== DIAGNOSTIC: CHECKING SPREADSHEET MATH SHEET ===');
const mathSheet = wb.Sheets['Snow - Math Calculations'];
if (mathSheet) {
  const mathData = sheetToArray(mathSheet);
  console.log('Snow - Math Calculations first 5 rows:');
  for (let r = 0; r < Math.min(5, mathData.length); r++) {
    console.log(`  Row ${r}: ${(mathData[r] || []).slice(0, 20).join(' | ')}`);
  }
} else {
  console.log('"Snow - Math Calculations" sheet not found');
}

// Check if there's a Snow - Engineering or similar sheet
const snowSheets = wb.SheetNames.filter(n => n.toLowerCase().includes('snow'));
console.log(`\nSnow-related sheets: ${snowSheets.join(', ')}`);

// ══════════════════════════════════════════════════════════════
// KEY INVESTIGATION: Why $280 in app vs $0 in spreadsheet?
// ══════════════════════════════════════════════════════════════
console.log('\n=== KEY INVESTIGATION ===');
console.log('Spreadsheet shows $- (zero/dash) for 22x50x12 with 30GL/90mph/AZ');
console.log(`App buckets 90mph → ${appBucketedWind}`);
console.log(`Truss spacing for [${appSnowCode}][${appFinalConfigKey}] = ${finalTrussSpacing}`);

if (finalTrussSpacing === 0) {
  console.log('\nHYPOTHESIS: Truss spacing = 0 means "Contact Engineering"');
  console.log('The app returns -1 for this case, which the UI may display as $280 erroneously');
  console.log('OR the app may be mapping -1 to some default price');
}

// Check what the spreadsheet ACTUALLY has for wind 90 specifically
console.log('\n--- What does the wind bucketing row say about 90? ---');
// Check raw data for column where input = 90
for (let c = 1; c < (changersData[0]?.length || 0); c++) {
  const inputVal = num(changersData[0][c]);
  if (inputVal >= 88 && inputVal <= 92) {
    console.log(`  Col ${c}: input=${changersData[0][c]}, bucket=${changersData[1]?.[c]}`);
  }
}

// Check ALL truss spacing entries for width 22 AFV
console.log('\n--- All truss spacing entries for width 22 AFV ---');
for (const [rowKey, cols] of Object.entries(trussSpacingMatrix)) {
  for (const [colKey, val] of Object.entries(cols)) {
    if (colKey.includes('22') && colKey.includes('AFV')) {
      console.log(`  [${rowKey}][${colKey}] = ${val}`);
    }
  }
}

// What if the spreadsheet treats 90mph as NO engineering needed?
console.log('\n--- Does 90mph even appear in the wind bucketing? ---');
const allWindInputs = Object.keys(windLoadBuckets).map(Number).sort((a,b) => a-b);
console.log(`  Wind inputs in bucketing table: ${allWindInputs.join(', ')}`);
const minWindInput = Math.min(...allWindInputs);
console.log(`  Minimum wind input: ${minWindInput}`);
if (WIND_MPH < minWindInput) {
  console.log(`\n  *** 90mph < ${minWindInput} (minimum in bucketing table) ***`);
  console.log('  *** The spreadsheet may treat values below the minimum as "no engineering needed" ***');
  console.log('  *** This means the spreadsheet shows $0 because 90mph is BELOW the threshold ***');
  console.log('  *** But the app\'s nearestBucket(90, [105,115,...]) returns 105, treating it AS IF 105mph ***');
  console.log('  *** THIS IS THE BUG: The app should return $0 when wind is below the minimum category ***');
}
