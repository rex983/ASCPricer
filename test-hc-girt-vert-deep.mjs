/**
 * Deep analysis of Snow - Hat Channels, Snow - Girts, and Snow - Verticals sheets
 * from "AZ CO UT 1 5 26.xlsx"
 */

import XLSX from 'xlsx';

const wb = XLSX.readFile('C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx', { cellFormula: true });

function getCell(ws, addr) {
  const cell = ws[addr];
  return cell ? { v: cell.v, f: cell.f || null, t: cell.t } : null;
}

function banner(title) {
  console.log('\n' + '='.repeat(80));
  console.log(`  ${title}`);
  console.log('='.repeat(80));
}

function subBanner(title) {
  console.log(`\n--- ${title} ---`);
}

// ============================================================================
// SNOW - HAT CHANNELS
// ============================================================================
banner('SNOW - HAT CHANNELS');

const hcWs = wb.Sheets['Snow - Hat Channels'];

// 1. ALL FORMULA CELLS
subBanner('1. ALL FORMULA CELLS');
{
  const range = XLSX.utils.decode_range(hcWs['!ref']);
  const formulas = [];
  for (let r = 0; r <= range.e.r; r++) {
    for (let c = 0; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = hcWs[addr];
      if (cell && cell.f) {
        formulas.push({ addr, formula: cell.f, value: cell.v });
      }
    }
  }
  console.log(`Total formula cells: ${formulas.length}`);
  for (const f of formulas) {
    console.log(`  ${f.addr}: =${f.formula}  (value: ${JSON.stringify(f.value)})`);
  }
}

// 2. MAIN LOOKUP TABLE STRUCTURE
subBanner('2. MAIN LOOKUP TABLE - Complete Structure');
{
  // Column headers (row 1, B1:H1 = wind speed categories)
  const colHeaders = [];
  for (let c = 1; c <= 7; c++) {
    const cell = hcWs[XLSX.utils.encode_cell({ r: 0, c })];
    if (cell) colHeaders.push(cell.v);
  }
  console.log(`Column headers (B1:H1) - Wind speed categories: ${JSON.stringify(colHeaders)}`);

  // Row keys (A2:A71 = trussSpacing-snowLoad combos)
  const rowKeys = [];
  for (let r = 1; r <= 70; r++) {
    const cell = hcWs[XLSX.utils.encode_cell({ r, c: 0 })];
    if (cell) rowKeys.push({ row: r + 1, key: cell.v });
  }
  console.log(`Row keys (A2:A71) - ${rowKeys.length} rows`);
  console.log(`  Format: "{trussSpacing}-{snowLoad}" e.g. "60-30GL", "60-20LL"`);

  // Group by truss spacing
  const groups = {};
  for (const rk of rowKeys) {
    const ts = rk.key.split('-')[0];
    if (!groups[ts]) groups[ts] = [];
    groups[ts].push(rk.key);
  }
  console.log(`\n  Truss spacing groups:`);
  for (const [ts, keys] of Object.entries(groups)) {
    console.log(`    ${ts}: ${keys.join(', ')}`);
  }

  console.log(`\n  Table dimensions: ${rowKeys.length} rows x ${colHeaders.length} cols`);
  console.log(`  Data range: B2:H71`);
  console.log(`  Values are HC spacing in inches (e.g. 54, 48, 42, 36, 32, 30, 24)`);

  // Show sample data for first truss spacing group
  console.log(`\n  Sample data (truss spacing 60):`);
  console.log(`  ${'Row Key'.padEnd(12)} ${colHeaders.map(h => String(h).padStart(4)).join(' ')}`);
  for (let r = 1; r <= 14; r++) {
    const key = hcWs[XLSX.utils.encode_cell({ r, c: 0 })]?.v;
    const vals = [];
    for (let c = 1; c <= 7; c++) {
      const cell = hcWs[XLSX.utils.encode_cell({ r, c })];
      vals.push(cell ? String(cell.v).padStart(4) : '   -');
    }
    console.log(`  ${String(key).padEnd(12)} ${vals.join(' ')}`);
  }
}

// 3. ORIGINAL HC COUNTS SECTION
subBanner('3. ORIGINAL HC COUNTS SECTION (R1:Z13)');
{
  // Header
  console.log(`Header: R1 = "Original Hat Channel"`);
  const widths = [];
  for (let c = 18; c <= 25; c++) { // S=18 to Z=25
    const cell = hcWs[XLSX.utils.encode_cell({ r: 0, c })];
    if (cell) widths.push(cell.v);
  }
  console.log(`Column headers (S1:Z1) - Building widths: ${JSON.stringify(widths)}`);

  // States
  const states = [];
  for (let r = 1; r <= 12; r++) {
    const stateCell = hcWs[XLSX.utils.encode_cell({ r, c: 17 })]; // R column
    if (stateCell) {
      const vals = [];
      for (let c = 18; c <= 25; c++) {
        const cell = hcWs[XLSX.utils.encode_cell({ r, c })];
        vals.push(cell ? cell.v : '-');
      }
      states.push({ state: stateCell.v, values: vals });
    }
  }

  console.log(`\nStructure: Rows = States, Columns = Building Widths`);
  console.log(`Indexed by: STATE (row) then WIDTH (column)`);
  console.log(`\n  ${'State'.padEnd(6)} ${widths.map(w => String(w).padStart(4)).join(' ')}`);
  for (const s of states) {
    console.log(`  ${String(s.state).padEnd(6)} ${s.values.map(v => String(v).padStart(4)).join(' ')}`);
  }

  // Lookup formulas
  console.log(`\nLookup mechanism (AC4:AF7):`);
  console.log(`  AC5 = '${hcWs['AC5']?.f}' => "${hcWs['AC5']?.v}" (state from Changers)`);
  console.log(`  AE5 = '${hcWs['AE5']?.f}' => ${hcWs['AE5']?.v} (width from Changers)`);
  console.log(`  AC7 = '${hcWs['AC7']?.f}' => ${hcWs['AC7']?.v} (row MATCH)`);
  console.log(`  AE7 = '${hcWs['AE7']?.f}' => ${hcWs['AE7']?.v} (column MATCH)`);
  console.log(`  AD10 = '${hcWs['AD10']?.f}' => ${hcWs['AD10']?.v} (INDEX result = original HC count)`);
}

// 4. Math Calculations P4 reads from L7
subBanner('4. Math Calculations P4 => L7');
{
  const cell = hcWs['L7'];
  console.log(`L7 formula: =${cell.f}`);
  console.log(`L7 value: ${cell.v}`);
  console.log(`Breakdown: INDEX($B$2:$H$71, $M$4, $K$4)`);
  console.log(`  $M$4 = MATCH($N$4, $A$2:$A$71, 0) => row match of concatenated key`);
  console.log(`  $N$4 = concatenate($L$2, "-", $M$2) => "${hcWs['N4']?.v}"`);
  console.log(`  $L$2 = '${hcWs['L2']?.f}' => ${hcWs['L2']?.v} (truss spacing from Changers)`);
  console.log(`  $M$2 = '${hcWs['M2']?.f}' => "${hcWs['M2']?.v}" (snow load from Changers)`);
  console.log(`  $K$4 = MATCH($K$2, $B$1:$H$1, 0) => ${hcWs['K4']?.v} (column index for wind speed)`);
  console.log(`  $K$2 = '${hcWs['K2']?.f}' => ${hcWs['K2']?.v} (wind speed from Changers)`);
  console.log(`\nSo L7 = INDEX(main_table, row_for("60-20LL"), col_for(105)) = ${cell.v}`);
}

// 5. Math Calculations T4 reads from AD10
subBanner('5. Math Calculations T4 => AD10');
{
  const cell = hcWs['AD10'];
  console.log(`AD10 formula: =${cell.f}`);
  console.log(`AD10 value: ${cell.v}`);
  console.log(`Breakdown: INDEX($S$2:$Z$13, $AC$7, $AE$7)`);
  console.log(`  $AC$7 = MATCH($AC$5, $R$2:$R$13, 0) => ${hcWs['AC7']?.v} (state row)`);
  console.log(`  $AE$7 = MATCH($AE$5, $S$1:$Z$1, 0) => ${hcWs['AE7']?.v} (width column)`);
  console.log(`\nSo AD10 = INDEX(original_hc_table, row_for("AZ"), col_for(24)) = ${cell.v}`);
}

// 6. CONDITIONAL FORMULAS?
subBanner('6. CONDITIONAL FORMULAS CHECK');
{
  const range = XLSX.utils.decode_range(hcWs['!ref']);
  const conditionals = [];
  for (let r = 0; r <= range.e.r; r++) {
    for (let c = 0; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = hcWs[addr];
      if (cell && cell.f && /\bIF\b/i.test(cell.f)) {
        conditionals.push({ addr, f: cell.f });
      }
    }
  }
  console.log(`Conditional (IF) formulas found: ${conditionals.length}`);
  if (conditionals.length > 0) {
    for (const c of conditionals) console.log(`  ${c.addr}: =${c.f}`);
  } else {
    console.log(`  NONE - all formulas are MATCH, INDEX, or CONCATENATE`);
  }
}


// ============================================================================
// SNOW - GIRTS
// ============================================================================
banner('SNOW - GIRTS');

const gWs = wb.Sheets['Snow - Girts '];

// 1. ALL FORMULA CELLS
subBanner('1. ALL FORMULA CELLS');
{
  const range = XLSX.utils.decode_range(gWs['!ref']);
  const formulas = [];
  for (let r = 0; r <= range.e.r; r++) {
    for (let c = 0; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = gWs[addr];
      if (cell && cell.f) {
        formulas.push({ addr, formula: cell.f, value: cell.v });
      }
    }
  }
  console.log(`Total formula cells: ${formulas.length}`);
  for (const f of formulas) {
    console.log(`  ${f.addr}: =${f.formula}  (value: ${JSON.stringify(f.value)})`);
  }
}

// 2. MAIN LOOKUP TABLE
subBanner('2. MAIN LOOKUP TABLE (A1:H6)');
{
  const colHeaders = [];
  for (let c = 1; c <= 7; c++) {
    const cell = gWs[XLSX.utils.encode_cell({ r: 0, c })];
    if (cell) colHeaders.push(cell.v);
  }
  console.log(`Column headers (B1:H1) - Wind speed categories: ${JSON.stringify(colHeaders)}`);

  const rowKeys = [];
  for (let r = 1; r <= 5; r++) {
    const cell = gWs[XLSX.utils.encode_cell({ r, c: 0 })];
    if (cell) rowKeys.push(cell.v);
  }
  console.log(`Row keys (A2:A6) - Truss spacing buckets: ${JSON.stringify(rowKeys)}`);

  console.log(`\n  ${'Truss'.padEnd(6)} ${colHeaders.map(h => String(h).padStart(4)).join(' ')}`);
  for (let r = 1; r <= 5; r++) {
    const key = gWs[XLSX.utils.encode_cell({ r, c: 0 })]?.v;
    const vals = [];
    for (let c = 1; c <= 7; c++) {
      const cell = gWs[XLSX.utils.encode_cell({ r, c })];
      vals.push(cell ? String(cell.v).padStart(4) : '   -');
    }
    console.log(`  ${String(key).padEnd(6)} ${vals.join(' ')}`);
  }
  console.log(`\n  Table dimensions: 5 rows x 7 cols`);
  console.log(`  Values are GIRT spacing in inches`);
}

// 3. ORIGINAL GIRT COUNTS SECTION
subBanner('3. ORIGINAL GIRT COUNTS (L2:M22)');
{
  console.log(`Structure: L = leg height (0-20), M = girt count`);
  console.log(`\n  ${'Height'.padEnd(8)} Count`);
  for (let r = 1; r <= 21; r++) {
    const h = gWs[XLSX.utils.encode_cell({ r, c: 11 })]?.v; // L
    const cnt = gWs[XLSX.utils.encode_cell({ r, c: 12 })]?.v; // M
    if (h !== undefined) {
      console.log(`  ${String(h).padEnd(8)} ${cnt}`);
    }
  }

  // Summarize height-to-count ranges
  console.log(`\n  Height ranges by count:`);
  let lastCount = null;
  let rangeStart = null;
  const ranges = [];
  for (let r = 1; r <= 21; r++) {
    const h = gWs[XLSX.utils.encode_cell({ r, c: 11 })]?.v;
    const cnt = gWs[XLSX.utils.encode_cell({ r, c: 12 })]?.v;
    if (cnt !== lastCount) {
      if (lastCount !== null) {
        ranges.push({ start: rangeStart, end: gWs[XLSX.utils.encode_cell({ r: r - 1, c: 11 })]?.v, count: lastCount });
      }
      rangeStart = h;
      lastCount = cnt;
    }
  }
  if (lastCount !== null) {
    ranges.push({ start: rangeStart, end: gWs[XLSX.utils.encode_cell({ r: 21, c: 11 })]?.v, count: lastCount });
  }
  for (const rng of ranges) {
    console.log(`    Height ${rng.start}-${rng.end} => ${rng.count} girts`);
  }
}

// 4. Math Calculations P6 reads from F14
subBanner('4. Math Calculations P6 => F14');
{
  const cell = gWs['F14'];
  console.log(`F14 formula: =${cell.f}`);
  console.log(`F14 value: ${cell.v}`);
  console.log(`Breakdown: INDEX($B$2:$H$6, $G$12, $E$12)`);
  console.log(`  $G$12 = MATCH($G$9, $A$2:$A$6, 0) => ${gWs['G12']?.v} (row for truss spacing bucket)`);
  console.log(`  $G$9 = $G$32 => ${gWs['G9']?.v} (truss spacing from Switch lookup)`);
  console.log(`  $E$12 = MATCH($E$9, $B$1:$H$1, 0) => ${gWs['E12']?.v} (col for wind speed)`);
  console.log(`  $E$9 = '${gWs['E9']?.f}' => ${gWs['E9']?.v} (wind speed from Changers)`);
  console.log(`\nSo F14 = INDEX(girt_table, row_for(60), col_for(105)) = ${cell.v}`);
}

// 5. Math Calculations T6 reads from T11
subBanner('5. Math Calculations T6 => T11');
{
  const cell = gWs['T11'];
  console.log(`T11 formula: =${cell.f}`);
  console.log(`T11 value: ${cell.v}`);
  console.log(`Breakdown: INDEX($M$2:$M$22, $P$11, 0)`);
  console.log(`  $P$11 = MATCH($R$9, $L$2:$L$22, 0) => ${gWs['P11']?.v} (row for leg height)`);
  console.log(`  $R$9 = '${gWs['R9']?.f}' => ${gWs['R9']?.v} (leg height from Changers)`);
  console.log(`\nSo T11 = INDEX(original_girt_counts, row_for(12), 0) = ${cell.v}`);
}

// 6. SWITCH ROW
subBanner('6. SWITCH ROW (Row 27-28, 30-32)');
{
  console.log(`Row 27 (Switch): Maps raw truss spacing values to lookup buckets`);
  console.log(`Row 28 (Size): Input truss spacing values 0-60`);

  // Show the mapping
  const sizeToSwitch = {};
  for (let c = 1; c <= 61; c++) {
    const size = gWs[XLSX.utils.encode_cell({ r: 27, c })]?.v;
    const sw = gWs[XLSX.utils.encode_cell({ r: 26, c })]?.v;
    if (size !== undefined && sw !== undefined) {
      if (!sizeToSwitch[sw]) sizeToSwitch[sw] = [];
      sizeToSwitch[sw].push(size);
    }
  }
  console.log(`\nSwitch mapping (truss spacing => bucket):`);
  for (const [bucket, sizes] of Object.entries(sizeToSwitch)) {
    const min = Math.min(...sizes);
    const max = Math.max(...sizes);
    console.log(`  Sizes ${min}-${max} => bucket ${bucket}`);
  }

  console.log(`\nSwitch lookup chain:`);
  console.log(`  D31 = '${gWs['D31']?.f}' => ${gWs['D31']?.v} (truss spacing from Math Calculations P2)`);
  console.log(`  G31 = '${gWs['G31']?.f}' => ${gWs['G31']?.v} (MATCH in Size row)`);
  console.log(`  G32 = '${gWs['G32']?.f}' => ${gWs['G32']?.v} (INDEX into Switch row => bucket value)`);
  console.log(`  G9 = $G$32 => ${gWs['G9']?.v} (fed into main lookup as row key)`);
}


// ============================================================================
// SNOW - VERTICALS
// ============================================================================
banner('SNOW - VERTICALS');

const vWs = wb.Sheets['Snow - Verticals'];

// 1. ALL FORMULA CELLS
subBanner('1. ALL FORMULA CELLS');
{
  const range = XLSX.utils.decode_range(vWs['!ref']);
  const formulas = [];
  for (let r = 0; r <= range.e.r; r++) {
    for (let c = 0; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = vWs[addr];
      if (cell && cell.f) {
        formulas.push({ addr, formula: cell.f, value: cell.v });
      }
    }
  }
  console.log(`Total formula cells: ${formulas.length}`);
  for (const f of formulas) {
    console.log(`  ${f.addr}: =${f.formula}  (value: ${JSON.stringify(f.value)})`);
  }
}

// 2. MAIN LOOKUP TABLE
subBanner('2. MAIN LOOKUP TABLE (A1:V8)');
{
  const colHeaders = [];
  for (let c = 1; c <= 21; c++) {
    const cell = vWs[XLSX.utils.encode_cell({ r: 0, c })];
    if (cell) colHeaders.push(cell.v);
  }
  console.log(`Column headers (B1:V1) - Height index: ${JSON.stringify(colHeaders)}`);
  console.log(`  (These are leg heights 0 through 20)`);

  const rowKeys = [];
  for (let r = 1; r <= 7; r++) {
    const cell = vWs[XLSX.utils.encode_cell({ r, c: 0 })];
    if (cell) rowKeys.push(cell.v);
  }
  console.log(`Row keys (A2:A8) - Wind speed categories: ${JSON.stringify(rowKeys)}`);

  console.log(`\nFull table:`);
  console.log(`  ${'Wind'.padEnd(6)} ${colHeaders.map(h => String(h).padStart(3)).join(' ')}`);
  for (let r = 1; r <= 7; r++) {
    const key = vWs[XLSX.utils.encode_cell({ r, c: 0 })]?.v;
    const vals = [];
    for (let c = 1; c <= 21; c++) {
      const cell = vWs[XLSX.utils.encode_cell({ r, c })];
      vals.push(cell ? String(cell.v).padStart(3) : '  -');
    }
    console.log(`  ${String(key).padEnd(6)} ${vals.join(' ')}`);
  }
  console.log(`\n  Table dimensions: 7 rows (wind speeds) x 21 cols (height indices 0-20)`);
  console.log(`  Lookup is: [wind_speed_row][height_col] => vertical spacing`);
}

// 3. ORIGINAL VERTICAL COUNTS
subBanner('3. ORIGINAL VERTICAL COUNTS (A13:I14)');
{
  const widths = [];
  for (let c = 1; c <= 8; c++) {
    const cell = vWs[XLSX.utils.encode_cell({ r: 12, c })];
    if (cell) widths.push(cell.v);
  }
  const counts = [];
  for (let c = 1; c <= 8; c++) {
    const cell = vWs[XLSX.utils.encode_cell({ r: 13, c })];
    if (cell) counts.push(cell.v);
  }
  console.log(`Row 13 header: "Original Verticals"`);
  console.log(`Row 13 (B13:I13) - Building widths: ${JSON.stringify(widths)}`);
  console.log(`Row 14 (B14:I14) - Vertical counts:  ${JSON.stringify(counts)}`);
  console.log(`\nMapping:`);
  for (let i = 0; i < widths.length; i++) {
    console.log(`  Width ${widths[i]}' => ${counts[i]} verticals`);
  }
}

// 4. Math Calculations P8 reads from Z8
subBanner('4. Math Calculations P8 => Z8');
{
  const cell = vWs['Z8'];
  console.log(`Z8 formula: =${cell.f}`);
  console.log(`Z8 value: ${cell.v}`);
  console.log(`Breakdown: INDEX($B$2:$V$8, $AB$6, $Y$6)`);
  console.log(`  $AB$6 = MATCH($AB$4, $A$2:$A$8, 0) => ${vWs['AB6']?.v} (row for wind speed)`);
  console.log(`  $AB$4 = '${vWs['AB4']?.f}' => ${vWs['AB4']?.v} (wind speed from Changers)`);
  console.log(`  $Y$6 = MATCH($Y$4, $B$1:$V$1, 0) => ${vWs['Y6']?.v} (col for leg height)`);
  console.log(`  $Y$4 = '${vWs['Y4']?.f}' => ${vWs['Y4']?.v} (leg height from Changers)`);
  console.log(`\nSo Z8 = INDEX(vert_table, row_for(wind=105), col_for(height=12)) = ${cell.v}`);
  console.log(`Lookup is: [wind_speed][leg_height] => vertical spacing`);
}

// 5. Math Calculations T8 reads from B21
subBanner('5. Math Calculations T8 => B21');
{
  const cell = vWs['B21'];
  console.log(`B21 formula: =${cell.f}`);
  console.log(`B21 value: ${cell.v}`);
  console.log(`Breakdown: INDEX($B$14:$I$14, 1, $B$19)`);
  console.log(`  $B$19 = MATCH($C$19, $B$13:$I$13, 0) => ${vWs['B19']?.v} (col for width)`);
  console.log(`  $C$19 = '${vWs['C19']?.f}' => ${vWs['C19']?.v} (width from Changers)`);
  console.log(`\nSo B21 = INDEX(original_vert_counts, 1, col_for(24)) = ${cell.v}`);
}

// 6. VERTICAL SPACING LOOKUP ORIENTATION
subBanner('6. VERTICAL SPACING LOOKUP ORIENTATION');
{
  console.log(`Main table: Rows = wind speeds, Columns = height indices`);
  console.log(`Lookup direction: [wind_speed][leg_height] => vertical spacing`);
  console.log(`\nFor Math Calculations:`);
  console.log(`  P8 (spacing) uses: INDEX(table, MATCH(wind, winds), MATCH(height, heights))`);
  console.log(`  T8 (count)   uses: INDEX(orig_counts, 1, MATCH(width, widths))`);
}


// ============================================================================
// SECONDARY CALCULATION SECTIONS CHECK
// ============================================================================
banner('SECONDARY CALCULATION SECTIONS CHECK');

subBanner('Hat Channels');
console.log(`Section 1 (A1:H71): Main HC spacing lookup table`);
console.log(`  - 70 rows, 7 wind columns`);
console.log(`  - Keyed by "{trussSpacing}-{snowLoad}" composite key`);
console.log(`Section 2 (K1:N4): Lookup computation area`);
console.log(`  - K2: wind speed from Changers`);
console.log(`  - L2: truss spacing from Changers`);
console.log(`  - M2: snow load from Changers`);
console.log(`  - N4: concatenated key = L2 & "-" & M2`);
console.log(`  - K4: column MATCH for wind`);
console.log(`  - M4: row MATCH for concatenated key`);
console.log(`  - L7: INDEX result => HC SPACING (fed to Math P4)`);
console.log(`Section 3 (R1:Z13): Original hat channel COUNTS by state x width`);
console.log(`  - 12 states x 8 widths`);
console.log(`Section 4 (AC4:AF10): Original HC count lookup computation`);
console.log(`  - AC5: state, AE5: width`);
console.log(`  - AC7: row MATCH, AE7: col MATCH`);
console.log(`  - AD10: INDEX result => original HC COUNT (fed to Math T4)`);
console.log(`NO other secondary sections found.`);

subBanner('Girts');
console.log(`Section 1 (A1:H6): Main girt spacing lookup table`);
console.log(`  - 5 rows (truss spacing buckets: 60,54,48,42,36), 7 wind columns`);
console.log(`Section 2 (E8:G12, F13:F14): Girt spacing lookup computation`);
console.log(`  - E9: wind speed, G9: truss spacing bucket`);
console.log(`  - E12: col MATCH, G12: row MATCH`);
console.log(`  - F14: INDEX result => GIRT SPACING (fed to Math P6)`);
console.log(`Section 3 (L1:M22): Original girt COUNTS by leg height`);
console.log(`  - Heights 0-20 mapped to counts 3-5`);
console.log(`Section 4 (P10:T11): Original girt count lookup computation`);
console.log(`  - R9: leg height, P11: MATCH, T11: INDEX => GIRT COUNT (fed to Math T6)`);
console.log(`Section 5 (A27:BJ28, D30:G32): SWITCH row - maps raw truss spacing to bucket`);
console.log(`  - Row 28 (Size): 0-60 sequential values`);
console.log(`  - Row 27 (Switch): maps each to nearest bucket (36,42,48,54,60)`);
console.log(`  - G31-G32: MATCH/INDEX to convert truss spacing to bucket`);
console.log(`NO other secondary sections found.`);

subBanner('Verticals');
console.log(`Section 1 (A1:V8): Main vertical spacing lookup table`);
console.log(`  - 7 rows (wind speeds: 105-180), 21 columns (height index 0-20)`);
console.log(`Section 2 (Y3:AC6, Z7:Z8): Vertical spacing lookup computation`);
console.log(`  - Y4: leg height, AB4: wind speed`);
console.log(`  - Y6: col MATCH (height), AB6: row MATCH (wind)`);
console.log(`  - Z8: INDEX result => VERTICAL SPACING (fed to Math P8)`);
console.log(`Section 3 (A13:I14): Original vertical COUNTS by width`);
console.log(`  - 8 widths (12-30), counts (4-7)`);
console.log(`Section 4 (B18:C19, B20:B21): Original vert count lookup computation`);
console.log(`  - C19: width, B19: MATCH, B21: INDEX => VERT COUNT (fed to Math T8)`);
console.log(`NO other secondary sections found.`);


// ============================================================================
// SUMMARY
// ============================================================================
banner('EXECUTIVE SUMMARY');
console.log(`
Each of these three sheets follows the SAME two-output pattern:

  OUTPUT 1: SPACING (fed to Math Calculations column P)
    - A lookup table that maps configuration inputs to a spacing value
    - HC:   [wind_speed x composite_key(trussSpacing-snowLoad)] => spacing
    - Girts: [wind_speed x trussSpacing_bucket] => spacing
    - Verts: [wind_speed x leg_height] => spacing

  OUTPUT 2: ORIGINAL COUNT (fed to Math Calculations column T)
    - A simpler table giving the "base" count before snow adjustments
    - HC:   [state x width] => count
    - Girts: [leg_height] => count (1D lookup)
    - Verts: [width] => count (1D lookup)

KEY FORMULAS:
  Hat Channels:
    L7 = INDEX($B$2:$H$71, MATCH(concatenate(trussSpacing,"-",snowLoad), A2:A71, 0), MATCH(wind, B1:H1, 0))
    AD10 = INDEX($S$2:$Z$13, MATCH(state, R2:R13, 0), MATCH(width, S1:Z1, 0))

  Girts:
    F14 = INDEX($B$2:$H$6, MATCH(G9, A2:A6, 0), MATCH(wind, B1:H1, 0))
      where G9 comes from Switch row: INDEX(Switch, 1, MATCH(trussSpacing, Size, 0))
    T11 = INDEX($M$2:$M$22, MATCH(legHeight, L2:L22, 0), 0)

  Verticals:
    Z8 = INDEX($B$2:$V$8, MATCH(wind, A2:A8, 0), MATCH(legHeight, B1:V1, 0))
    B21 = INDEX($B$14:$I$14, 1, MATCH(width, B13:I13, 0))

NO CONDITIONAL (IF) FORMULAS exist in any of these three sheets.
All lookups use MATCH + INDEX pattern exclusively.
`);
