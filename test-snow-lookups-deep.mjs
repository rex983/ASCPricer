/**
 * Deep analysis of ALL snow lookup sheets in the ASC Pricing spreadsheet.
 * Exhaustively documents every sheet's structure, keys, headers, and values.
 */

import XLSX from 'xlsx';

const wb = XLSX.readFile('C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx');

function getActualDataExtent(data) {
  let lastRow = 0;
  let maxCol = 0;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (row && row.some(v => v !== '' && v !== undefined && v !== null)) {
      lastRow = i;
      for (let j = row.length - 1; j >= 0; j--) {
        if (row[j] !== '' && row[j] !== undefined && row[j] !== null) {
          maxCol = Math.max(maxCol, j + 1);
          break;
        }
      }
    }
  }
  return { rows: lastRow + 1, cols: maxCol };
}

function printSeparator(title) {
  console.log('\n' + '='.repeat(80));
  console.log(`  ${title}`);
  console.log('='.repeat(80));
}

// ============================================================================
// 1. SNOW - TRUSS SPACING
// ============================================================================
printSeparator('1. SNOW - TRUSS SPACING');

{
  const ws = wb.Sheets['Snow - Truss Spacing'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const extent = getActualDataExtent(data);

  const headers = data[0].filter(h => h !== '');
  const rowKeys = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== '' && data[i][0] !== undefined) rowKeys.push(data[i][0]);
  }

  console.log(`\nDimensions: ${extent.rows} rows × ${extent.cols} cols`);
  console.log(`Column headers: ${headers.length}`);
  console.log(`Row keys: ${rowKeys.length}`);

  // Parse header format
  const headerParts = new Set();
  const enclosures = new Set();
  const winds = new Set();
  const widths = new Set();
  const roofStyles = new Set();

  for (const h of headers) {
    const parts = h.split('-');
    if (parts.length === 4) {
      enclosures.add(parts[0]);
      winds.add(parseInt(parts[1]));
      widths.add(parseInt(parts[2]));
      roofStyles.add(parts[3]);
    }
  }

  console.log(`\nHeader format: {enclosure}-{windSpeed}-{width}-{roofStyle}`);
  console.log(`  Enclosures: ${[...enclosures].sort().join(', ')}`);
  console.log(`  Wind speeds: ${[...winds].sort((a, b) => a - b).join(', ')}`);
  console.log(`  Widths: ${[...widths].sort((a, b) => a - b).join(', ')}`);
  console.log(`  Roof styles: ${[...roofStyles].sort().join(', ')}`);

  // Parse row key format
  const rowPrefixes = new Set();
  const snowLoads = new Set();
  for (const k of rowKeys) {
    const parts = k.split('-');
    rowPrefixes.add(parts[0]);
    snowLoads.add(parts[1]);
  }

  console.log(`\nRow key format: {sizePrefix}-{snowLoad}`);
  console.log(`  Size prefixes: ${[...rowPrefixes].sort().join(', ')}`);
  console.log(`  Snow loads: ${[...snowLoads].join(', ')}`);
  console.log(`\nAll row keys:`);
  console.log(`  ${rowKeys.join(', ')}`);

  console.log(`\nAll column headers:`);
  // Group by width for readability
  for (const w of [...widths].sort((a, b) => a - b)) {
    const matching = headers.filter(h => h.includes(`-${w}-`));
    console.log(`  Width ${w}: ${matching.join(', ')}`);
  }

  // Sample lookups
  console.log(`\nSample lookups:`);
  const headerIdx = {};
  data[0].forEach((h, i) => { if (h) headerIdx[h] = i; });
  const rowIdx = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) rowIdx[data[i][0]] = i;
  }

  const samples = [
    ['T-30GL', 'E-105-12-STD'],
    ['T-70GL', 'E-180-30-AFV'],
    ['M-50GL', 'O-140-24-STD'],
    ['S-20LL', 'E-130-20-AFV'],
  ];
  for (const [rk, ck] of samples) {
    if (rowIdx[rk] !== undefined && headerIdx[ck] !== undefined) {
      console.log(`  [${rk}][${ck}] = ${data[rowIdx[rk]][headerIdx[ck]]}`);
    }
  }

  // Value ranges
  const allValues = [];
  for (let i = 1; i <= rowKeys.length; i++) {
    for (let j = 1; j < data[0].length; j++) {
      if (data[i] && data[i][j] !== '' && data[i][j] !== undefined) {
        allValues.push(data[i][j]);
      }
    }
  }
  const uniqueVals = [...new Set(allValues)].sort((a, b) => a - b);
  console.log(`\nUnique spacing values: ${uniqueVals.join(', ')}`);
  console.log(`  (These are truss spacings in inches)`);
}

// ============================================================================
// 2. SNOW - TRUSSES
// ============================================================================
printSeparator('2. SNOW - TRUSSES');

{
  const ws = wb.Sheets['Snow - Trusses '];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const extent = getActualDataExtent(data);

  const headers = data[0].filter(h => h !== '');
  const rowKeys = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== '' && data[i][0] !== undefined) rowKeys.push(data[i][0]);
  }

  console.log(`\nDimensions: ${extent.rows} rows × ${extent.cols} cols`);
  console.log(`Column headers: ${headers.length}`);
  console.log(`Row keys: ${rowKeys.length} (1 to ${rowKeys[rowKeys.length - 1]})`);

  // Parse header format
  const widthsSet = new Set();
  const statesSet = new Set();
  for (const h of headers) {
    const parts = h.split('-');
    if (parts.length === 2) {
      widthsSet.add(parseInt(parts[0]));
      statesSet.add(parts[1]);
    }
  }

  console.log(`\nHeader format: {width}-{stateCode}`);
  console.log(`  Widths: ${[...widthsSet].sort((a, b) => a - b).join(', ')}`);
  console.log(`  States: ${[...statesSet].sort().join(', ')}`);

  console.log(`\nAll column headers:`);
  for (const state of [...statesSet].sort()) {
    const matching = headers.filter(h => h.endsWith(`-${state}`));
    console.log(`  ${state}: ${matching.join(', ')}`);
  }

  console.log(`\nRow keys are building LENGTH in feet (1 to 100)`);

  // Sample lookups
  console.log(`\nSample lookups (length -> truss count by width-state):`);
  const headerIdx = {};
  data[0].forEach((h, i) => { if (h) headerIdx[h] = i; });

  const samples = [
    [50, '24-AZ'],
    [50, '24-OH'],
    [30, '12-TX'],
    [100, '30-CA'],
  ];
  for (const [len, key] of samples) {
    if (headerIdx[key] !== undefined) {
      console.log(`  Length ${len}, ${key} = ${data[len][headerIdx[key]]} trusses`);
    }
  }

  // Value ranges
  const allValues = new Set();
  for (let i = 1; i <= 100; i++) {
    for (let j = 1; j <= headers.length; j++) {
      if (data[i] && data[i][j] !== '' && data[i][j] !== undefined) {
        allValues.add(data[i][j]);
      }
    }
  }
  console.log(`\nUnique truss count values: ${[...allValues].sort((a, b) => a - b).join(', ')}`);
}

// ============================================================================
// 3. SNOW - HAT CHANNELS
// ============================================================================
printSeparator('3. SNOW - HAT CHANNELS');

{
  const ws = wb.Sheets['Snow - Hat Channels'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const extent = getActualDataExtent(data);

  // Main lookup table: columns A-H (indices 0-7)
  const headers = data[0].slice(1, 8); // wind speeds
  const rowKeys = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== '' && data[i][0] !== undefined) rowKeys.push(data[i][0]);
  }

  console.log(`\nDimensions: ${extent.rows} rows × ${extent.cols} cols`);
  console.log(`\n--- MAIN LOOKUP TABLE (cols A-H) ---`);
  console.log(`Column headers (wind speeds): ${headers.join(', ')}`);
  console.log(`Row keys: ${rowKeys.length}`);
  console.log(`All row keys: ${rowKeys.join(', ')}`);

  // Parse row key format
  const trussSpacings = new Set();
  const snowLoadCodes = new Set();
  for (const k of rowKeys) {
    const parts = k.split('-');
    trussSpacings.add(parseInt(parts[0]));
    snowLoadCodes.add(parts[1]);
  }

  console.log(`\nRow key format: {trussSpacing}-{snowLoadCode}`);
  console.log(`  Truss spacings: ${[...trussSpacings].sort((a, b) => a - b).join(', ')}`);
  console.log(`  Snow load codes: ${[...snowLoadCodes].join(', ')}`);

  // Sample lookups
  console.log(`\nSample lookups:`);
  const headerIdx = { 105: 1, 115: 2, 130: 3, 140: 4, 155: 5, 165: 6, 180: 7 };
  const rowIdx = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) rowIdx[data[i][0]] = i;
  }

  const samples = ['60-30GL', '60-70GL', '36-20LL', '42-50GL', '48-61LL'];
  for (const rk of samples) {
    if (rowIdx[rk] !== undefined) {
      const row = data[rowIdx[rk]];
      console.log(`  ${rk}: 105=${row[1]}, 115=${row[2]}, 130=${row[3]}, 140=${row[4]}, 155=${row[5]}, 165=${row[6]}, 180=${row[7]}`);
    }
  }

  // Original hat channel counts (right side of sheet)
  console.log(`\n--- ORIGINAL HAT CHANNEL COUNTS (cols R-Y) ---`);
  console.log(`Header row: "Original Hat Channel" with widths: ${data[0].slice(17, 25).join(', ')}`);
  console.log(`\nState -> original HC count by width:`);
  for (let i = 1; i <= 12; i++) {
    const state = data[i][9]; // column J
    if (state && state !== '') {
      const counts = data[i].slice(17, 25);
      console.log(`  ${state}: ${counts.join(', ')}`);
    }
  }

  // Unique HC spacing values
  const allValues = new Set();
  for (const rk of rowKeys) {
    const r = rowIdx[rk];
    for (let j = 1; j <= 7; j++) {
      if (data[r][j] !== '' && data[r][j] !== undefined) allValues.add(data[r][j]);
    }
  }
  console.log(`\nUnique hat channel spacing values: ${[...allValues].sort((a, b) => a - b).join(', ')}`);
}

// ============================================================================
// 4. SNOW - GIRTS
// ============================================================================
printSeparator('4. SNOW - GIRTS');

{
  const ws = wb.Sheets['Snow - Girts '];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const extent = getActualDataExtent(data);

  console.log(`\nDimensions: ${extent.rows} rows × ${extent.cols} cols`);

  // Main lookup table
  console.log(`\n--- MAIN LOOKUP TABLE (cols A-H) ---`);
  console.log(`Column headers: "Girt Spacing", ${data[0].slice(1, 8).join(', ')}`);

  const mainRows = [];
  for (let i = 1; i <= 5; i++) {
    if (data[i][0] !== '' && data[i][0] !== undefined && typeof data[i][0] === 'number') {
      mainRows.push(data[i][0]);
      console.log(`  Height ${data[i][0]}: ${data[i].slice(1, 8).join(', ')}`);
    }
  }
  console.log(`Row keys (leg heights): ${mainRows.join(', ')}`);

  // Original girt counts
  console.log(`\n--- ORIGINAL GIRT COUNTS (cols K-L) ---`);
  console.log(`Format: height -> girt count`);
  for (let i = 1; i <= 25; i++) {
    const height = data[i][11]; // col L (index 11)
    const count = data[i][12]; // col M (index 12)
    if (height !== '' && height !== undefined && count !== '' && count !== undefined) {
      console.log(`  Height ${height}: ${count} girts`);
    }
  }

  // Switch and Size rows
  console.log(`\n--- SWITCH ROW (girt spacing threshold by size) ---`);
  const switchRow = data[26];
  const sizeRow = data[27];
  console.log(`Sizes: 0 to ${sizeRow[sizeRow.length - 1]}`);
  // Find transitions in Switch row
  let lastVal = null;
  const transitions = [];
  for (let i = 1; i < switchRow.length; i++) {
    if (switchRow[i] !== lastVal && switchRow[i] !== '' && switchRow[i] !== undefined) {
      transitions.push({ size: sizeRow[i], spacing: switchRow[i] });
      lastVal = switchRow[i];
    }
  }
  console.log(`Transitions (size -> girt spacing switch point):`);
  for (const t of transitions) {
    console.log(`  Size ${t.size}+: girt spacing = ${t.spacing}`);
  }

  // Lookup helper rows
  console.log(`\n--- LOOKUP HELPER AREA (cols D-G, rows 7-13) ---`);
  for (let i = 7; i <= 15; i++) {
    const row = data[i];
    if (row && row.some(v => v !== '')) {
      const label = row.slice(3, 8).filter(v => v !== '').join(' | ');
      if (label) console.log(`  Row ${i}: ${label}`);
    }
  }
}

// ============================================================================
// 5. SNOW - VERTICALS
// ============================================================================
printSeparator('5. SNOW - VERTICALS');

{
  const ws = wb.Sheets['Snow - Verticals'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const extent = getActualDataExtent(data);

  console.log(`\nDimensions: ${extent.rows} rows × ${extent.cols} cols`);

  // Main lookup table
  console.log(`\n--- MAIN LOOKUP TABLE ---`);
  console.log(`Column header row: "Spacing", then leg height indices: ${data[0].slice(1, 22).join(', ')}`);
  console.log(`Row keys (wind speeds):`);

  for (let i = 1; i <= 7; i++) {
    if (data[i][0] !== '' && typeof data[i][0] === 'number') {
      console.log(`  Wind ${data[i][0]}: ${data[i].slice(1, 22).join(', ')}`);
    }
  }

  console.log(`\nLookup: row=windSpeed, col=legHeightIndex -> vertical spacing`);
  console.log(`Leg height indices 0-20 correspond to specific leg heights`);

  // Original verticals
  console.log(`\n--- ORIGINAL VERTICALS (row 12-13) ---`);
  console.log(`Widths: ${data[12].slice(1, 9).join(', ')}`);
  console.log(`Counts: ${data[13].slice(1, 9).join(', ')}`);

  // Unique values
  const allValues = new Set();
  for (let i = 1; i <= 7; i++) {
    for (let j = 1; j <= 21; j++) {
      if (data[i][j] !== '' && data[i][j] !== undefined) allValues.add(data[i][j]);
    }
  }
  console.log(`\nUnique vertical spacing values: ${[...allValues].sort((a, b) => a - b).join(', ')}`);

  // Helper/match section on right
  console.log(`\n--- HELPER / MATCH SECTION (right side) ---`);
  for (let i = 2; i <= 8; i++) {
    const right = data[i].slice(22);
    if (right && right.some(v => v !== '')) {
      console.log(`  Row ${i}: ${right.filter(v => v !== '').join(' | ')}`);
    }
  }
}

// ============================================================================
// 6. SNOW - DIAGONAL BRACING
// ============================================================================
printSeparator('6. SNOW - DIAGONAL BRACING');

{
  const ws = wb.Sheets['Snow - Diagonal Bracing'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const extent = getActualDataExtent(data);

  console.log(`\nDimensions: ${extent.rows} rows × ${extent.cols} cols`);
  console.log(`\nThis sheet is NOT a simple lookup table — it's a CALCULATION sheet.`);

  console.log(`\n--- CALCULATION LOGIC ---`);
  console.log(`1. Check if building is partially enclosed: row 2, col F = ${data[2][6]}`);
  console.log(`2. What is the wind load: row 4, col F = ${data[4][6]}`);
  console.log(`3. Required wind load for DB: row 5, col F = ${data[5][6]}`);
  console.log(`4. Subtract (wind - required): row 6, col F = ${data[6][6]}`);
  console.log(`5. If negative, doesn't need DB: row 7, col F = ${data[7][6]}`);
  console.log(`6. Price per DB: row 8, col K = ${data[8][10]}`);

  console.log(`\n--- REQUIRED WIND BY STATE (row 12-13) ---`);
  const states = data[12].slice(1);
  const requiredWinds = data[13].slice(1);
  const stateWindMap = {};
  for (let i = 0; i < states.length; i++) {
    if (states[i] !== '' && states[i] !== undefined) {
      stateWindMap[states[i]] = requiredWinds[i];
    }
  }
  console.log(`State -> Required wind speed for diagonal bracing:`);
  // Group by unique values
  const byWind = {};
  for (const [state, wind] of Object.entries(stateWindMap)) {
    if (!byWind[wind]) byWind[wind] = [];
    byWind[wind].push(state);
  }
  for (const [wind, stateList] of Object.entries(byWind)) {
    console.log(`  ${wind} MPH: ${stateList.join(', ')}`);
  }

  console.log(`\n--- OPEN/ENCLOSED LOGIC (rows 21-27) ---`);
  console.log(`Calculation determines DB based on:`);
  console.log(`  - Building type (open vs fully enclosed)`);
  console.log(`  - Building width (50' and under vs over 50')`);
  console.log(`  - Leg height affects count`);

  // Leg height section
  console.log(`\n--- LEG HEIGHT -> DB COUNT (row 2-3, col N-O) ---`);
  console.log(`Leg height: ${data[2][14]}, factor: ${data[4][14]}`);

  // Pricing
  console.log(`\n--- PRICING ---`);
  console.log(`Rows 22-26 show pricing logic:`);
  console.log(`  Open + <=50': count=${data[22][16]}, price/ea=${data[22][18]}, total=${data[22][20]}`);
  console.log(`  Enclosed + <=50': count=${data[23][16]}, price/ea=${data[23][18]}, total=${data[23][20]}`);
  console.log(`  Open + >50': count=${data[25][16]}, price/ea=${data[25][18]}, total=${data[25][20]}`);
  console.log(`  Enclosed + >50': count=${data[26][16]}, price/ea=${data[26][18]}, total=${data[26][20]}`);
  console.log(`  Final DB price: ${data[27][20]}`);
}

// ============================================================================
// 7. SNOW - CHANGERS (bonus - mapping/translation sheet)
// ============================================================================
printSeparator('7. SNOW - CHANGERS (mapping/translation sheet)');

{
  const ws = wb.Sheets['Snow - Changers'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const extent = getActualDataExtent(data);

  console.log(`\nDimensions: ${extent.rows} rows × ${extent.cols} cols`);

  console.log(`\n--- WIND LOAD CHANGERS (rows 0-5) ---`);
  console.log(`Maps position index to wind load used for roof calculations`);
  console.log(`Row 0 (Wind Load): positions 0-${data[0].length - 2}`);
  console.log(`Row 1 (Roof): all values = ${[...new Set(data[1].slice(1).filter(v => v !== ''))].join(', ')}`);
  console.log(`Wind changer logic: input ${data[3][4]}, changer ${data[4][4]}, actual used ${data[5][4]}`);

  console.log(`\n--- SNOW LOAD CODES (rows 8-9) ---`);
  console.log(`Snow Load labels: ${data[8].slice(1, 15).join(', ')}`);
  console.log(`Value codes:      ${data[9].slice(1, 15).join(', ')}`);

  console.log(`\n--- HC CHART MAPPING (rows 16-17) ---`);
  console.log(`Maps position index to "Based on TS" (truss spacing):`);
  const uniqueTS = [...new Set(data[17].slice(1).filter(v => v !== ''))];
  console.log(`Unique truss spacings in HC chart: ${uniqueTS.join(', ')}`);

  console.log(`\n--- LEG HEIGHT MAPPING (rows 25-27) ---`);
  console.log(`Leg height position -> size category (S/M/T) and feet used`);
  console.log(`LH positions: ${data[25].slice(1, 22).join(', ')}`);
  console.log(`Categories:   ${data[26].slice(1, 22).join(', ')}`);
  console.log(`Feet used:    ${data[27].slice(1, 22).join(', ')}`);

  console.log(`\n--- ROOF STYLE MAPPING (rows 41-46) ---`);
  console.log(`Symbols: ${data[42].slice(0, 5).join(', ')}`);

  console.log(`\n--- WIDTH MAPPING (rows 48-49) ---`);
  console.log(`Position: ${data[48].slice(1, 29).join(', ')}`);
  console.log(`Width:    ${data[49].slice(1, 29).join(', ')}`);

  console.log(`\n--- STATE MAPPING (rows 56-58) ---`);
  console.log(`Flyer groups: ${data[56].filter(v => v !== '').join(', ')}`);
  console.log(`Full names:   ${data[57].slice(1).filter(v => v !== '').join(', ')}`);
  console.log(`State codes:  ${data[58].slice(1).filter(v => v !== '').join(', ')}`);

  console.log(`\n--- TRUSS PRICING BY STATE & WIDTH (rows 59-66) ---`);
  const widthRows = [59, 60, 61, 62, 63, 64, 65, 66];
  for (const r of widthRows) {
    const label = data[r][0];
    const prices = [...new Set(data[r].slice(1).filter(v => v !== ''))];
    console.log(`  ${label}: unique prices = ${prices.join(', ')}`);
  }

  console.log(`\n--- PER-UNIT PRICING BY STATE (rows 67-69) ---`);
  console.log(`  Pie truss: ${[...new Set(data[67].slice(1).filter(v => v !== ''))].join(', ')}`);
  console.log(`  Channel $/ft: ${[...new Set(data[68].slice(1).filter(v => v !== ''))].join(', ')}`);
  console.log(`  Tubing $/ft: ${[...new Set(data[69].slice(1).filter(v => v !== ''))].join(', ')}`);

  console.log(`\n--- SNOW LOAD FULL NAMES (row 79) ---`);
  console.log(`Codes: ${data[79].slice(1, 15).join(', ')}`);
  console.log(`Names: ${data[79].slice(19, 30).filter(v => v !== '').join(', ')}`);
}

// ============================================================================
// 8. SNOW - MATH CALCULATIONS (bonus - formula logic)
// ============================================================================
printSeparator('8. SNOW - MATH CALCULATIONS (formula logic)');

{
  const ws = wb.Sheets['Snow - Math Calculations'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  console.log(`\n--- TRUSS SPACING MATH (rows 0-6) ---`);
  console.log(`  Length: ${data[1][3]} ft`);
  console.log(`  To inches: ${data[2][3]}`);
  console.log(`  Divide by spacing: ${data[3][3]}`);
  console.log(`  Round up: ${data[4][3]}`);
  console.log(`  Add 1: ${data[5][3]} (trusses required)`);
  console.log(`  Extra trusses: ${data[6][3]} (subtract original trusses)`);

  console.log(`\n--- HAT CHANNEL MATH (rows 8-16) ---`);
  console.log(`  Width: ${data[9][3]} ft`);
  console.log(`  Finding bar size: ${data[10][3]}`);
  console.log(`  Half: ${data[11][3]}`);
  console.log(`  To inches: ${data[12][3]}`);
  console.log(`  Divide by spacing: ${data[13][3]}`);
  console.log(`  Round up: ${data[14][3]}`);
  console.log(`  Total: ${data[15][3]} (HC required per bay)`);
  console.log(`  Extra: ${data[16][3]} (subtract original HC)`);

  console.log(`\n--- GIRT MATH (rows 18-25) ---`);
  console.log(`  Leg height: ${data[19][3]} ft`);
  console.log(`  To inches: ${data[20][3]}`);
  console.log(`  Needed girts: ${data[21][3]}`);
  console.log(`  Round up: ${data[22][3]}`);
  console.log(`  Add 1: ${data[23][3]}`);
  console.log(`  Extra girts: ${data[24][3]}`);
  console.log(`  Only V sides: ${data[25][3]}`);
  console.log(`  FE or O: ${data[19][7]} | Yes/No: ${data[20][7]}`);
  console.log(`  IF E or O: ${data[21][7]} | IF Y/N: ${data[22][7]}`);
  console.log(`  Are girts needed: ${data[23][7]} (1=Y / 0=N)`);

  console.log(`\n--- VERTICAL MATH (rows 27-34) ---`);
  console.log(`  Width: ${data[28][3]} ft`);
  console.log(`  Leg height: ${data[29][3]} ft`);
  console.log(`  Width to inches: ${data[30][3]}`);
  console.log(`  Divide by spacing: ${data[31][3]}`);
  console.log(`  Round up: ${data[32][3]}`);
  console.log(`  Add 1: ${data[33][3]}`);
  console.log(`  Subtract OV: ${data[34][3]} extra verticals per side`);

  console.log(`\n--- REQUIRED SPACINGS (current building) ---`);
  console.log(`  Truss spacing required: ${data[1][15]}"`);
  console.log(`  HC spacing required: ${data[3][15]}"`);
  console.log(`  Girt spacing required: ${data[5][15]}"`);
  console.log(`  Vertical spacing required: ${data[7][15]}"`);

  console.log(`\n--- ORIGINAL COUNTS ---`);
  console.log(`  Original trusses: ${data[1][19]}`);
  console.log(`  Original hat channels: ${data[3][19]}`);
  console.log(`  Original girts: ${data[5][19]}`);
  console.log(`  Original verticals: ${data[7][19]}`);

  console.log(`\n--- PRICING FORMULAS (rows 11-29) ---`);
  console.log(`  Extra trusses: ${data[12][15]}, truss price: ${data[13][15]}`);
  console.log(`  Extra leg height from 6': ${data[15][15]} -> price/truss: ${data[16][15]}`);
  console.log(`  Total truss price: ${data[18][15]}`);
  console.log(`  Extra verticals: ${data[12][20]}, peak height: ${data[13][20]}`);
  console.log(`  Tubing used: ${data[13][25]}, price/vertical: ${data[17][20]}`);
  console.log(`  Extra channels: ${data[24][15]}, price/ft: ${data[25][15]}`);
  console.log(`  Channel length: ${data[26][15]}, price/channel: ${data[27][15]}`);
  console.log(`  Girt: width=${data[24][19]}, length=${data[25][19]}, perimeter=${data[26][19]}`);
  console.log(`  Tubing price/ft: ${data[28][19]}`);
}

// ============================================================================
// SUMMARY
// ============================================================================
printSeparator('SUMMARY OF ALL SNOW LOAD CODES');

console.log(`
SNOW LOAD CODES (Ground Loads - GL):
  30GL, 40GL, 50GL, 60GL, 70GL, 80GL, 90GL

SNOW LOAD CODES (Roof/Live Loads - LL):
  20LL, 27LL, 34LL, 41LL, 47LL, 54LL, 61LL

SIZE PREFIXES (in Truss Spacing sheet):
  T = Tall (leg heights 10-20)
  M = Medium (leg heights 7-9)
  S = Short/Standard (leg heights 0-6)

WIND SPEEDS: 105, 115, 130, 140, 155, 165, 180

BUILDING WIDTHS: 12, 18, 20, 22, 24, 26, 28, 30

ROOF STYLES: STD (Standard), AFV (A-Frame Vertical)

ENCLOSURE TYPES: E (Enclosed), O (Open)

STATES: OH, MI, TX, IL, WI, MO, AZ, NM, CA, NV, NY, PA

TRUSS SPACINGS (possible values in inches):
  12, 18, 24, 30, 32, 36, 40, 42, 48, 54, 60

LOOKUP FLOW:
  1. Truss Spacing: [sizePrefix-snowLoad] x [enclosure-wind-width-roofStyle] -> spacing in inches
  2. Trusses: [length] x [width-state] -> original truss count
  3. Hat Channels: [trussSpacing-snowLoad] x [windSpeed] -> HC spacing in inches
  4. Girts: [legHeight] x [windSpeed] -> girt spacing in inches
  5. Verticals: [windSpeed] x [legHeightIndex] -> vertical spacing in inches
  6. Diagonal Bracing: calculation based on wind vs required wind by state

MATH FLOW (for each component):
  1. Get required spacing from lookup
  2. Convert building dimension to inches
  3. Divide by required spacing
  4. Round up, add 1
  5. Subtract original count = extra needed
  6. Multiply by per-unit price = cost
`);

console.log('\nScript complete.');
