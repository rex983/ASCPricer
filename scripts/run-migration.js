/**
 * Run a migration against Supabase.
 * Usage: node scripts/run-migration.js <database-password> [migration-filename]
 *   migration-filename defaults to the most recent file in supabase/migrations/.
 *
 * Get your DB password from: Supabase Dashboard > Settings > Database > Connection string
 */
const { Client } = require("pg");
const fs = require("fs");
const path = require("path");

const password = process.argv[2];
const filenameArg = process.argv[3];
if (!password) {
  console.error("Usage: node scripts/run-migration.js <database-password> [migration-filename]");
  process.exit(1);
}

const migrationsDir = path.join(__dirname, "../supabase/migrations");
const filename =
  filenameArg ||
  fs.readdirSync(migrationsDir).filter((f) => f.endsWith(".sql")).sort().pop();

const sqlPath = path.join(migrationsDir, filename);
if (!fs.existsSync(sqlPath)) {
  console.error(`Migration file not found: ${sqlPath}`);
  process.exit(1);
}

console.log(`Running migration: ${filename}`);
const sql = fs.readFileSync(sqlPath, "utf8");

const client = new Client({
  host: "aws-0-us-east-1.pooler.supabase.com",
  port: 5432,
  database: "postgres",
  user: "postgres.xockuiyvxijuzlwlsfbu",
  password,
  ssl: { rejectUnauthorized: false },
});

(async () => {
  await client.connect();
  console.log("Connected. Running migration...");
  await client.query(sql);
  console.log("Migration completed successfully!");
  await client.end();
})().catch((e) => {
  console.error("Migration failed:", e.message);
  process.exit(1);
});
