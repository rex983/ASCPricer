import XLSX from 'xlsx';
import { readFileSync } from 'fs';

const wb = XLSX.readFile('C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx', {
  cellFormula: true,
  cellStyles: true,
});

const ws = wb.Sheets['Snow - Math Calculations'];
if (!ws) {
  console.log('Sheet not found. Available:', wb.SheetNames);
  process.exit(1);
}

const range = XLSX.utils.decode_range(ws['!ref']);

// ============================================================
// HELPER: Collect all cells
// ============================================================
const allCells = [];
for (let r = range.s.r; r <= range.e.r; r++) {
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ws[addr];
    if (cell) {
      allCells.push({ addr, r, c, cell });
    }
  }
}

// ============================================================
// 1. CROSS-SHEET REFERENCE ANALYSIS
// ============================================================
console.log('='.repeat(100));
console.log('SNOW - MATH CALCULATIONS: EXHAUSTIVE ANALYSIS');
console.log('='.repeat(100));

console.log('\n' + '='.repeat(100));
console.log('SECTION 1: ALL CROSS-SHEET REFERENCES');
console.log('='.repeat(100));

const inboundRefs = new Map(); // sheet -> [{cell, formula}]
const outboundConsumers = new Map(); // We'll check other sheets

// Find all formulas referencing other sheets
const formulaCells = allCells.filter(c => c.cell.f);
const crossSheetPattern = /'([^']+)'!(\$?[A-Z]+\$?\d+)/g;

for (const { addr, cell } of formulaCells) {
  const formula = cell.f;
  let match;
  crossSheetPattern.lastIndex = 0;
  while ((match = crossSheetPattern.exec(formula)) !== null) {
    const sheet = match[1];
    const ref = match[2];
    if (!inboundRefs.has(sheet)) inboundRefs.set(sheet, []);
    inboundRefs.get(sheet).push({ localCell: addr, remoteRef: ref, formula });
  }
}

console.log('\n--- INBOUND: Other sheets that feed INTO this sheet ---');
for (const [sheet, refs] of [...inboundRefs.entries()].sort()) {
  console.log(`\n  Sheet: "${sheet}" (${refs.length} references)`);
  const unique = new Map();
  for (const r of refs) {
    const key = `${sheet}!${r.remoteRef}`;
    if (!unique.has(key)) unique.set(key, []);
    unique.get(key).push(r.localCell);
  }
  for (const [ref, cells] of unique) {
    console.log(`    ${ref} -> used in: ${cells.join(', ')}`);
  }
}

// Check other sheets for references TO this sheet
console.log('\n--- OUTBOUND: Other sheets that READ FROM this sheet ---');
const thisSheetNames = ['Snow - Math Calculations', 'Snow - Math'];
for (const sheetName of wb.SheetNames) {
  if (sheetName === 'Snow - Math Calculations') continue;
  const otherWs = wb.Sheets[sheetName];
  const otherRange = otherWs['!ref'] ? XLSX.utils.decode_range(otherWs['!ref']) : null;
  if (!otherRange) continue;

  const refs = [];
  for (let r = otherRange.s.r; r <= otherRange.e.r; r++) {
    for (let c = otherRange.s.c; c <= otherRange.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = otherWs[addr];
      if (cell && cell.f) {
        for (const name of thisSheetNames) {
          if (cell.f.includes(name)) {
            refs.push({ addr, formula: cell.f });
          }
        }
      }
    }
  }
  if (refs.length > 0) {
    console.log(`\n  Sheet: "${sheetName}" (${refs.length} references to this sheet)`);
    for (const r of refs) {
      console.log(`    ${sheetName}!${r.addr}: =${r.formula.substring(0, 120)}${r.formula.length > 120 ? '...' : ''}`);
    }
  }
}

// ============================================================
// 2. SECTION-BY-SECTION DETAILED BREAKDOWN
// ============================================================
console.log('\n\n' + '='.repeat(100));
console.log('SECTION 2: DETAILED FORMULA CHAINS BY COMPONENT');
console.log('='.repeat(100));

// Section boundaries based on merged headers
const sections = [
  { name: 'TRUSSES (Extra Trusses Needed)', startRow: 0, endRow: 7, cols: 'A-H' },
  { name: 'TRUSSES - Lookup Data', startRow: 1, endRow: 8, cols: 'M-T' },
  { name: 'HAT CHANNELS (Extra Hat Channels Needed)', startRow: 8, endRow: 17, cols: 'A-H' },
  { name: 'HAT CHANNELS - AFV Check', startRow: 10, endRow: 13, cols: 'F-I' },
  { name: 'GIRTS (Extra Girts Needed - V Sides Only)', startRow: 18, endRow: 26, cols: 'A-I' },
  { name: 'GIRTS - Enclosed Check', startRow: 19, endRow: 26, cols: 'F-I' },
  { name: 'VERTICALS (Extra Verticals Needed)', startRow: 27, endRow: 35, cols: 'A-I' },
  { name: 'TRUSS PRICING', startRow: 11, endRow: 22, cols: 'M-P' },
  { name: 'VERTICAL PRICING', startRow: 11, endRow: 22, cols: 'R-V' },
  { name: 'HAT CHANNEL PRICING', startRow: 23, endRow: 29, cols: 'M-P' },
  { name: 'GIRT PRICING', startRow: 23, endRow: 30, cols: 'R-V' },
  { name: 'OUTPUT SUMMARY', startRow: 2, endRow: 7, cols: 'W-AF' },
  { name: 'TOTAL MATERIAL OUTPUT', startRow: 11, endRow: 23, cols: 'W-AF' },
];

for (const section of sections) {
  console.log(`\n${'─'.repeat(80)}`);
  console.log(`  ${section.name} (Rows ${section.startRow + 1}-${section.endRow + 1}, Cols ${section.cols})`);
  console.log(`${'─'.repeat(80)}`);

  const colStart = section.cols.split('-')[0].charCodeAt(0) - 65;
  const colEnd = section.cols.split('-')[1].charCodeAt(0) - 65;

  for (let r = section.startRow; r <= section.endRow; r++) {
    const rowCells = [];
    for (let c = colStart; c <= colEnd; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell && (cell.v !== undefined || cell.f)) {
        rowCells.push({ addr, cell });
      }
    }
    if (rowCells.length === 0) continue;

    for (const { addr, cell } of rowCells) {
      const label = cell.t === 's' && !cell.f ? `[LABEL] "${cell.v}"` : '';
      const formula = cell.f ? `=${cell.f}` : '';
      const value = cell.v !== undefined && cell.t !== 's' ? `VALUE=${JSON.stringify(cell.v)}` : cell.t === 's' && cell.f ? `VALUE="${cell.v}"` : '';
      if (label || formula) {
        console.log(`  ${addr}: ${label}${formula} ${value}`.trimEnd());
      }
    }
  }
}

// ============================================================
// 3. COMPLETE CALCULATION FLOW DIAGRAMS
// ============================================================
console.log('\n\n' + '='.repeat(100));
console.log('SECTION 3: COMPLETE CALCULATION FLOW (INPUT -> OUTPUT)');
console.log('='.repeat(100));

console.log(`
╔══════════════════════════════════════════════════════════════════════════════════╗
║  COMPONENT 1: EXTRA TRUSSES                                                    ║
╠══════════════════════════════════════════════════════════════════════════════════╣
║                                                                                 ║
║  INPUTS:                                                                        ║
║    D2  = 'Quote Sheet'!N10         (Building Length = ${ws['D2']?.v})                       ║
║    P2  = 'Snow - Truss Spacing'!F54 (Required Spacing = ${ws['P2']?.v}" o.c.)               ║
║    T2  = 'Snow - Trusses'!CV11     (Original Trusses = ${ws['T2']?.v})                     ║
║                                                                                 ║
║  CALCULATION:                                                                   ║
║    D3  = D2 * 12                   (Length in inches = ${ws['D3']?.v})                    ║
║    D4  = D3 / P2                   (Bays at spacing = ${ws['D4']?.v})                      ║
║    D5  = ROUNDUP(D4, 0)            (Rounded up = ${ws['D5']?.v})                           ║
║    D6  = D5 + 1                    (Trusses needed = ${ws['D6']?.v})                       ║
║    D7  = D6 - T2                   (Extra = ${ws['D7']?.v})                                ║
║    G2  = IF(D7 < 0, 0, 1)          (Neg guard = ${ws['G2']?.v})                            ║
║    G7  = D7 * G2                   (Final extra = ${ws['G7']?.v})                          ║
║                                                                                 ║
║  OUTPUT:                                                                        ║
║    G7 -> P13 (extra trusses count for pricing)                                  ║
║    G7 -> X5  (output summary, guarded by P2=0 check)                            ║
║                                                                                 ║
║  ZERO CONDITIONS:                                                               ║
║    - D7 < 0 (original trusses >= calculated needed) => G7 = 0                  ║
║    - P2 = 0 (no spacing found) => X5 = "Please Contact Engineering"            ║
╚══════════════════════════════════════════════════════════════════════════════════╝

╔══════════════════════════════════════════════════════════════════════════════════╗
║  COMPONENT 2: EXTRA HAT CHANNELS                                               ║
╠══════════════════════════════════════════════════════════════════════════════════╣
║                                                                                 ║
║  INPUTS:                                                                        ║
║    D10 = 'Snow - Changers'!D54     (Width = ${ws['D10']?.v})                               ║
║    P4  = 'Snow - Hat Channels'!L7  (Required Spacing = ${ws['P4']?.v}" o.c.)               ║
║    T4  = 'Snow - Hat Channels'!AD10 (Original Hat Ch = ${ws['T4']?.v})                     ║
║    H11 = 'Snow - Changers'!D47     (Roof Type = "${ws['H11']?.v}")                       ║
║                                                                                 ║
║  CALCULATION:                                                                   ║
║    D11 = D10 + 2                   (Bar size = ${ws['D11']?.v})                            ║
║    D12 = D11 / 2                   (Half = ${ws['D12']?.v})                                ║
║    D13 = D12 * 12                  (Inches = ${ws['D13']?.v})                             ║
║    D14 = D13 / P4                  (Channels at spacing = ${ws['D14']?.v?.toFixed(4)})              ║
║    D15 = ROUNDUP(D14, 0)           (Rounded = ${ws['D15']?.v})                             ║
║    D16 = (D15 + 1) * 2             (Both sides +1 each = ${ws['D16']?.v})                  ║
║    H12 = IF(H11="AFV", 1, 0)       (Is vertical roof? = ${ws['H12']?.v})                  ║
║    D17 = (D16 - T4) * H12          (Extra channels = ${ws['D17']?.v})                      ║
║    H13 = IF(D17 < 0, 0, 1)         (Neg guard = ${ws['H13']?.v})                          ║
║    G17 = D17 * H13                 (Final extra = ${ws['G17']?.v})                         ║
║                                                                                 ║
║  KEY LOGIC: Hat channels ONLY apply to AFV (A-Frame Vertical) roofs!            ║
║    If roof is NOT AFV, H12=0, so D17=0 always => no extra channels              ║
║                                                                                 ║
║  OUTPUT:                                                                        ║
║    G17 -> P25 (extra channels count for pricing)                                ║
║                                                                                 ║
║  ZERO CONDITIONS:                                                               ║
║    - Roof is NOT "AFV" => H12=0 => D17=0 => G17=0                              ║
║    - D17 < 0 (original >= needed) => H13=0 => G17=0                            ║
╚══════════════════════════════════════════════════════════════════════════════════╝

╔══════════════════════════════════════════════════════════════════════════════════╗
║  COMPONENT 3: EXTRA GIRTS (V-sides only)                                       ║
╠══════════════════════════════════════════════════════════════════════════════════╣
║                                                                                 ║
║  INPUTS:                                                                        ║
║    D20 = 'Snow - Changers'!D29     (Leg Height = ${ws['D20']?.v})                          ║
║    P6  = 'Snow - Girts'!F14        (Required Spacing = ${ws['P6']?.v}" o.c.)               ║
║    T6  = 'Snow - Girts'!T11        (Original Girts = ${ws['T6']?.v})                      ║
║    H20 = 'Pricing - Changers'!D66  (Frame type = "${ws['H20']?.v}")                      ║
║    H21 = 'Pricing - Changers'!U69  (Some flag = "${ws['H21']?.v}")                       ║
║                                                                                 ║
║  CALCULATION:                                                                   ║
║    D21 = D20 * 12                  (LH in inches = ${ws['D21']?.v})                       ║
║    D22 = D21 / P6                  (Girts at spacing = ${ws['D22']?.v})                    ║
║    D23 = ROUNDUP(D22, 0)           (Rounded = ${ws['D23']?.v})                             ║
║    D24 = D23 + 1                   (Plus 1 = ${ws['D24']?.v})                              ║
║    D25 = D24 - T6                  (Extra raw = ${ws['D25']?.v})                           ║
║                                                                                 ║
║  GIRT GATE (are girts needed?):                                                 ║
║    H22 = IF(H20="E", 1, 0)         (Is enclosed? = ${ws['H22']?.v})                       ║
║    H23 = IFS(H21=0,0; "No",0; "Yes",1) (Flag check = ${ws['H23']?.v})                    ║
║    H24 = H22 * H23                 (Both must be true = ${ws['H24']?.v})                   ║
║                                                                                 ║
║    D26 = D25 * H24                 (Gated extra = ${ws['D26']?.v})                         ║
║    H25 = IF(D26 < 0, 0, 1)         (Neg guard = ${ws['H25']?.v})                          ║
║    H26 = D26 * H25                 (Final extra = ${ws['H26']?.v})                         ║
║                                                                                 ║
║  KEY LOGIC:                                                                     ║
║    Girts only added when BOTH: frame="E" (Enclosed) AND flag="Yes"              ║
║    Either condition false => H24=0 => D26=0 => no extra girts                   ║
║                                                                                 ║
║  OUTPUT:                                                                        ║
║    H26 -> T28 (feeds girt pricing total feet calc)                              ║
║                                                                                 ║
║  ZERO CONDITIONS:                                                               ║
║    - Frame type is not "E" => H22=0 => H24=0 => no girts                       ║
║    - Flag H21 is "No" or 0 => H23=0 => H24=0 => no girts                       ║
║    - D26 < 0 => H25=0 => H26=0                                                 ║
╚══════════════════════════════════════════════════════════════════════════════════╝

╔══════════════════════════════════════════════════════════════════════════════════╗
║  COMPONENT 4: EXTRA VERTICALS                                                  ║
╠══════════════════════════════════════════════════════════════════════════════════╣
║                                                                                 ║
║  INPUTS:                                                                        ║
║    D29 = 'Snow - Changers'!D54     (Width = ${ws['D29']?.v})                               ║
║    D30 = 'Snow - Changers'!D29     (Leg Height = ${ws['D30']?.v})                          ║
║    P8  = 'Snow - Verticals'!Z8     (Required Spacing = ${ws['P8']?.v}" o.c.)               ║
║    T8  = 'Snow - Verticals'!B21    (Original Verticals = ${ws['T8']?.v})                   ║
║    I34 = IF('Quote Sheet'!F16="Enclosed Ends",1,0) = ${ws['I34']?.v}                   ║
║    I35 = I34 * 'Quote Sheet'!R16   (Enclosed ends * qty = ${ws['I35']?.v})                 ║
║                                                                                 ║
║  CALCULATION:                                                                   ║
║    D31 = D29 * 12                  (Width inches = ${ws['D31']?.v})                       ║
║    D32 = D31 / P8                  (Verticals at spacing = ${ws['D32']?.v})                ║
║    D33 = ROUNDUP(D32, 0)           (Rounded = ${ws['D33']?.v})                             ║
║    D34 = D33 + 1                   (Plus 1 = ${ws['D34']?.v})                              ║
║    D35 = (D34 - T8) * I35          (Extra * end multiplier = ${ws['D35']?.v})              ║
║    H31 = IF(D35 < 0, 0, 1)         (Neg guard = ${ws['H31']?.v})                          ║
║    H32 = H31 * D35                 (Final extra = ${ws['H32']?.v})                         ║
║                                                                                 ║
║  KEY LOGIC:                                                                     ║
║    Verticals multiply by I35 which is 0 when NOT "Enclosed Ends"                ║
║    When enclosed, I35 = number of enclosed end walls (from Quote Sheet!R16)      ║
║                                                                                 ║
║  OUTPUT:                                                                        ║
║    H32 -> U13 (feeds vertical pricing)                                          ║
║    H32 -> AA5 (output summary, guarded by P8=0 check)                           ║
║                                                                                 ║
║  ZERO CONDITIONS:                                                               ║
║    - Not "Enclosed Ends" => I34=0 => I35=0 => D35=0 => H32=0                   ║
║    - D35 < 0 => H31=0 => H32=0                                                 ║
║    - P8 = 0 => AA5 = "Please Contact Engineering for Info"                      ║
╚══════════════════════════════════════════════════════════════════════════════════╝
`);

// ============================================================
// 4. PRICING CALCULATION FLOWS
// ============================================================
console.log('='.repeat(100));
console.log('SECTION 4: PRICING CALCULATION FLOWS');
console.log('='.repeat(100));

console.log(`
╔══════════════════════════════════════════════════════════════════════════════════╗
║  TRUSS PRICING (Cols M-P, Rows 12-22)                                          ║
╠══════════════════════════════════════════════════════════════════════════════════╣
║                                                                                 ║
║  P13 = G7                          (Extra trusses = ${ws['P13']?.v})                       ║
║  P14 = 'Snow - Changers'!G76       (Truss unit price = $${ws['P14']?.v})                  ║
║  P15 = P13 * P14                   (Base truss cost = $${ws['P15']?.v})                   ║
║                                                                                 ║
║  Leg Height Adder:                                                              ║
║  P16 = 'Snow - Changers'!J72       (Leg ht price/ft = $${ws['P16']?.v})                   ║
║  P17 = D20 - 6                     (Extra ft above 6' = ${ws['P17']?.v})                   ║
║  P18 = IF(P17 < 0, 0, 1)           (Neg guard = ${ws['P18']?.v})                          ║
║  P19 = Z14 = 'Snow-Changers'!G32   (Tubing used = ${ws['P19']?.v})                        ║
║  P20 = P19 * P16                   (Extra LH cost/truss = $${ws['P20']?.v})               ║
║  P21 = P13 * P20                   (Total LH adder = $${ws['P21']?.v})                    ║
║                                                                                 ║
║  P22 = P15 + P21                   (TOTAL TRUSS PRICE = $${ws['P22']?.v})                 ║
║                                                                                 ║
║  NOTE: P19 uses Z14 which is tubing size from Snow-Changers, NOT P17*P18        ║
║  The P18 guard is calculated but NOT used in the pricing chain!                 ║
║  (P17/P18 seem vestigial or for display only)                                   ║
║                                                                                 ║
║  OUTPUT: P22 -> X6 (guarded by P2=0), also -> AE14                             ║
╚══════════════════════════════════════════════════════════════════════════════════╝

╔══════════════════════════════════════════════════════════════════════════════════╗
║  HAT CHANNEL PRICING (Cols M-P, Rows 24-29)                                    ║
╠══════════════════════════════════════════════════════════════════════════════════╣
║                                                                                 ║
║  P25 = G17                         (Extra channels = ${ws['P25']?.v})                      ║
║  P26 = 'Snow - Changers'!J75       (Price per foot = $${ws['P26']?.v})                    ║
║  P27 = D2 + 1                      (Length + 1 = ${ws['P27']?.v})                          ║
║  P28 = P27 * P26                   (Price per channel = $${ws['P28']?.v})                ║
║  P29 = P25 * P28                   (TOTAL CHANNEL PRICE = $${ws['P29']?.v})               ║
║                                                                                 ║
║  OUTPUT: P29 -> AE15                                                            ║
╚══════════════════════════════════════════════════════════════════════════════════╝

╔══════════════════════════════════════════════════════════════════════════════════╗
║  VERTICAL PRICING (Cols R-V, Rows 12-22)                                        ║
╠══════════════════════════════════════════════════════════════════════════════════╣
║                                                                                 ║
║  U13 = H32                         (Extra verticals = ${ws['U13']?.v})                     ║
║  U14 = ((D10/2)*3)/12              (Peak height = ${ws['U14']?.v})                         ║
║  U15 = ROUNDUP(U14, 0)             (Rounded peak = ${ws['U15']?.v})                        ║
║  U16 = D20 + U15                   (Total vert height = ${ws['U16']?.v})                   ║
║  U17 = 'Snow - Changers'!J76       (Tubing $/ft = $${ws['U17']?.v})                      ║
║  U18 = U16 * U17                   (Price per vertical = $${ws['U18']?.v})                ║
║  U19 = U18 * U13                   (Base vert price = $${ws['U19']?.v})                   ║
║                                                                                 ║
║  U20 = COMPLEX TIERED IF:                                                       ║
║    Width (Quote Sheet!R10) determines multiplier:                                ║
║      13-15 ft: U19 * 2             (double)                                     ║
║      16-18 ft: (U18 * 2.5) * U13   (2.5x per vert)                             ║
║      19-20 ft: (U18 * 3) * U13     (3x per vert)                               ║
║      else:     U19                  (1x, no multiplier)                          ║
║    Current value = $${ws['U20']?.v}                                                       ║
║                                                                                 ║
║  OUTPUT: U20 -> AA6, AE17                                                       ║
╚══════════════════════════════════════════════════════════════════════════════════╝

╔══════════════════════════════════════════════════════════════════════════════════╗
║  GIRT PRICING (Cols R-V, Rows 24-30)                                            ║
╠══════════════════════════════════════════════════════════════════════════════════╣
║                                                                                 ║
║  T25 = D10                         (Width = ${ws['T25']?.v})                               ║
║  T26 = D2                          (Length = ${ws['T26']?.v})                              ║
║  V25 = 'Pricing - Changers'!U64    (Sides V = ${ws['V25']?.v})                            ║
║  V26 = 'Pricing - Changers'!U65    (Ends V = ${ws['V26']?.v})                             ║
║  V27 = 'Pricing - Changers'!H65    (Sides QTY = ${ws['V27']?.v})                          ║
║  V28 = 'Pricing - Changers'!H66    (Ends QTY = ${ws['V28']?.v})                           ║
║  V29 = V25 * V27 * T26             (Sides perimeter = ${ws['V29']?.v})                     ║
║  V30 = V26 * V28 * T25             (Ends perimeter = ${ws['V30']?.v})                      ║
║  T27 = V29 + V30                   (Total perimeter = ${ws['T27']?.v})                     ║
║  T28 = T27 * H26                   (Total feet = ${ws['T28']?.v})                          ║
║  T29 = 'Snow - Changers'!J76       (Tubing $/ft = $${ws['T29']?.v})                      ║
║  T30 = T28 * T29                   (TOTAL GIRT PRICE = $${ws['T30']?.v})                  ║
║                                                                                 ║
║  OUTPUT: T30 -> AE16                                                            ║
╚══════════════════════════════════════════════════════════════════════════════════╝
`);

// ============================================================
// 5. FINAL OUTPUT / TOTAL MATERIAL SECTION
// ============================================================
console.log('='.repeat(100));
console.log('SECTION 5: FINAL OUTPUT SUMMARY');
console.log('='.repeat(100));

console.log(`
╔══════════════════════════════════════════════════════════════════════════════════╗
║  OUTPUT CELLS (Cols W-AF)                                                       ║
╠══════════════════════════════════════════════════════════════════════════════════╣
║                                                                                 ║
║  "Making Trusses not Error Symbol" (Row 3 header):                              ║
║                                                                                 ║
║  X5  = IF(P2=0, "Please Contact Engineering", G7)                               ║
║        => Trusses count = ${ws['X5']?.v}                                                   ║
║  AA5 = IF(P8=0, "Please Contact Engineering for Info", H32)                     ║
║        => Verticals count = ${ws['AA5']?.v}                                                ║
║  AC5 = 'Pricing - Anchors'!I58     (Anchors = ${ws['AC5']?.v})                             ║
║                                                                                 ║
║  X6  = IF(P2=0, "Please Contact Engineering for Info", P22)                     ║
║        => Truss Price = $${ws['X6']?.v}                                                   ║
║  AA6 = IF(P8=0, "Please Contact Engineering for Info", U20)                     ║
║        => Vertical Price = $${ws['AA6']?.v}                                               ║
║                                                                                 ║
║  X7  = IF(X5>=1, 'Snow-Changers'!G76 * X5, 0)                                  ║
║        => TRUSS CHARGE (alt calc) = $${ws['X7']?.v}                                       ║
║                                                                                 ║
║  TOTAL MATERIAL SECTION (Rows 12-20):                                           ║
║                                                                                 ║
║  Spacing Values:                                                                ║
║    AC14 = P2  (Truss spacing = ${ws['AC14']?.v})                                           ║
║    AC15 = P4  (Channel spacing = ${ws['AC15']?.v})                                         ║
║    AC16 = P6  (Girt spacing = ${ws['AC16']?.v})                                            ║
║    AC17 = P8  (Vertical spacing = ${ws['AC17']?.v})                                        ║
║    AC19 = AC14 * AC15 * AC16 * AC17 = ${ws['AC19']?.v}                         ║
║                                                                                 ║
║  Component Totals:                                                              ║
║    AE14 = X6  (Truss total = $${ws['AE14']?.v})                                          ║
║    AE15 = P29 (Channel total = $${ws['AE15']?.v})                                         ║
║    AE16 = T30 (Girt total = $${ws['AE16']?.v})                                            ║
║    AE17 = U20 (Vertical total = $${ws['AE17']?.v})                                        ║
║    AE19 = SUM(AE14:AE17) = $${ws['AE19']?.v}                                             ║
║                                                                                 ║
║  ★ FINAL OUTPUT ★                                                               ║
║  AD20 = IF(OR(AC19=0, AC14<18), "Contact Engineering", AE19)                    ║
║       = $${ws['AD20']?.v}                                                                 ║
║                                                                                 ║
║  Uprights Count:                                                                ║
║    Z13 = D20                       (Leg height = ${ws['Z13']?.v})                          ║
║    Z14 = 'Snow - Changers'!G32     (Tubing = ${ws['Z14']?.v})                              ║
║    Z21 = D6 * 2                    (Truss uprights = ${ws['Z21']?.v})                      ║
║    Z22 = D34 * 2                   (Vertical uprights = ${ws['Z22']?.v})                   ║
║    Z23 = Z21 + Z22                 (Total uprights = ${ws['Z23']?.v})                      ║
╚══════════════════════════════════════════════════════════════════════════════════╝
`);

// ============================================================
// 6. ALL IF/CONDITIONAL FORMULAS
// ============================================================
console.log('='.repeat(100));
console.log('SECTION 6: ALL CONDITIONAL LOGIC (IF/IFS statements)');
console.log('='.repeat(100));

const conditionalCells = formulaCells.filter(c =>
  /\b(IF|IFS|OR|AND)\b/i.test(c.cell.f)
);

for (const { addr, cell } of conditionalCells) {
  console.log(`\n  ${addr}: =${cell.f}`);
  console.log(`    Current value: ${JSON.stringify(cell.v)}`);

  // Analyze the condition
  if (cell.f.includes('Contact Engineering')) {
    console.log('    ⚠ CAN PRODUCE "Contact Engineering" OUTPUT');
  }
  if (/IF\([^,]+<\s*0/.test(cell.f)) {
    console.log('    → Negative guard: returns 0 when input is negative');
  }
  if (/IF\([^,]+=\s*0/.test(cell.f)) {
    console.log('    → Zero guard: returns error string when spacing is 0');
  }
}

// ============================================================
// 7. ZERO RESULT ANALYSIS
// ============================================================
console.log('\n\n' + '='.repeat(100));
console.log('SECTION 7: COMPLETE $0 / "CONTACT ENGINEERING" ANALYSIS');
console.log('='.repeat(100));

console.log(`
ALL PATHS THAT LEAD TO $0 RESULTS:
══════════════════════════════════

TRUSSES ($0 when):
  1. Extra trusses D7 < 0 (original count >= snow-required count)
     => G2=0, G7=0, P13=0, P15=0, P21=0, P22=0, X6=0, AE14=0

  2. Truss spacing P2 = 0 (lookup returned 0 / not found)
     => X5 = "Please Contact Engineering"
     => X6 = "Please Contact Engineering for Info"
     => Also triggers AD20 "Contact Engineering" (via AC19=0)

HAT CHANNELS ($0 when):
  1. Roof is NOT "AFV" => H12=0 => D17=0 => G17=0 => P25=0 => P29=0
  2. Extra channels D17 < 0 => H13=0 => G17=0 => P25=0 => P29=0

GIRTS ($0 when):
  1. Frame type H20 != "E" => H22=0 => H24=0 => D26=0 => H26=0 => T28=0 => T30=0
  2. Flag H21 = "No" or 0 => H23=0 => H24=0 (same chain)
  3. Extra girts D26 < 0 => H25=0 => H26=0 => T28=0 => T30=0
  4. All V-panel counts (V25, V26) = 0 => V29=0, V30=0 => T27=0 => T28=0

VERTICALS ($0 when):
  1. NOT "Enclosed Ends" => I34=0 => I35=0 => D35=0 => H32=0 => U13=0
  2. Extra verticals D35 < 0 => H31=0 => H32=0 => U13=0
  3. Vertical spacing P8=0 => AA5/AA6 = "Please Contact Engineering for Info"

OVERALL TOTAL AD20 = "Contact Engineering" when:
  1. AC19 = 0 (ANY spacing value is 0 => product is 0)
     => If P2=0 OR P4=0 OR P6=0 OR P8=0 => "Contact Engineering"
  2. AC14 < 18 (truss spacing < 18" o.c.)
     => "Contact Engineering"
  Otherwise: AD20 = AE19 = SUM of all 4 component prices

COMMON SCENARIO FOR ALL-ZERO:
  A non-snow-load region or configuration where all 4 snow spacing lookups
  return values >= the standard spacing => no extra components needed => $0.
  This is the NORMAL result for light snow areas.
`);

// ============================================================
// 8. EVERY FORMULA IN THE SHEET (raw dump)
// ============================================================
console.log('='.repeat(100));
console.log('SECTION 8: COMPLETE RAW FORMULA LIST (every formula cell)');
console.log('='.repeat(100));

for (const { addr, cell } of formulaCells.sort((a, b) => a.r - b.r || a.c - b.c)) {
  const v = cell.v !== undefined ? ` => ${JSON.stringify(cell.v)}` : '';
  console.log(`  ${addr}: =${cell.f}${v}`);
}

console.log(`\n\nTotal formula cells: ${formulaCells.length}`);
console.log(`Total non-empty cells: ${allCells.length}`);

// ============================================================
// 9. DEPENDENCY GRAPH SUMMARY
// ============================================================
console.log('\n' + '='.repeat(100));
console.log('SECTION 9: SHEET DEPENDENCY GRAPH');
console.log('='.repeat(100));

const sheetsIn = [...inboundRefs.keys()].sort();
console.log(`\nSheets that FEED INTO "Snow - Math Calculations":`);
for (const s of sheetsIn) {
  console.log(`  ← ${s} (${inboundRefs.get(s).length} refs)`);
}

console.log(`\nSheets that READ FROM "Snow - Math Calculations":`);
// Re-scan for outbound
for (const sheetName of wb.SheetNames) {
  if (sheetName === 'Snow - Math Calculations') continue;
  const otherWs = wb.Sheets[sheetName];
  const otherRange = otherWs['!ref'] ? XLSX.utils.decode_range(otherWs['!ref']) : null;
  if (!otherRange) continue;
  let count = 0;
  for (let r = otherRange.s.r; r <= otherRange.e.r; r++) {
    for (let c = otherRange.s.c; c <= otherRange.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = otherWs[addr];
      if (cell?.f && (cell.f.includes('Snow - Math Calculations') || cell.f.includes('Snow - Math'))) {
        count++;
      }
    }
  }
  if (count > 0) {
    console.log(`  → ${sheetName} (${count} refs)`);
  }
}

console.log('\n' + '='.repeat(100));
console.log('ANALYSIS COMPLETE');
console.log('='.repeat(100));
