const XLSX = require("xlsx");
const fs = require("fs");

const filePath = "C:/Users/Redir/Downloads/AZ CO UT 1 5 26.xlsx";
const sheetNames = [
  "Snow - Truss Spacing",
  "Snow - Trusses",
  "Snow - Hat Channels",
  "Snow - Girts",
  "Snow - Verticals",
  "Snow - Diagonal Bracing",
  "Snow Load Breakdown",
  "Pricing - Changers",
];

const wb = XLSX.readFile(filePath);

function findSheet(name) {
  if (wb.SheetNames.includes(name)) return name;
  const trimmed = wb.SheetNames.find(s => s.trim() === name.trim());
  if (trimmed) return trimmed;
  return null;
}

for (const name of sheetNames) {
  const actualName = findSheet(name);
  const safeName = name.replace(/[^a-zA-Z0-9]/g, '_');
  const outFile = `C:/Users/Redir/asc-pricing/dump-${safeName}.csv`;

  if (!actualName) {
    fs.writeFileSync(outFile, `SHEET NOT FOUND: ${name}\n`);
    console.log(`NOT FOUND: ${name}`);
    continue;
  }

  const ws = wb.Sheets[actualName];
  const csv = XLSX.utils.sheet_to_csv(ws, { blankrows: false });
  fs.writeFileSync(outFile, csv);
  const lines = csv.split('\n').filter(l => l.trim()).length;
  console.log(`Wrote ${outFile} (${lines} non-empty rows)`);
}
