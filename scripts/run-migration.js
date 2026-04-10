/**
 * Run pending migration against Supabase.
 * Usage: node scripts/run-migration.js <database-password>
 *
 * Get your DB password from: Supabase Dashboard > Settings > Database > Connection string
 */
const { Client } = require("pg");
const fs = require("fs");
const path = require("path");

const password = process.argv[2];
if (!password) {
  console.error("Usage: node scripts/run-migration.js <database-password>");
  process.exit(1);
}

const sql = fs.readFileSync(
  path.join(__dirname, "../supabase/migrations/20260409000000_create_sales_reps.sql"),
  "utf8"
);

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
