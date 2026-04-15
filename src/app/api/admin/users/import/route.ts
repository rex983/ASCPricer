import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { logAudit } from "@/lib/audit";

const VALID_ROLES = ["admin", "manager", "sales_rep", "viewer"];
const VALID_OFFICES = ["Harbor", "Marion"];
const ALLOWED_DOMAIN = "bigbuildingsdirect.com";
const MAX_FILE_BYTES = 1_000_000; // 1 MB
const MAX_ROWS = 1000;
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

interface CsvRow {
  name: string;
  email: string;
  role?: string;
  office?: string;
}

function parseCsv(text: string): CsvRow[] {
  const lines = text.split(/\r?\n/).filter((l) => l.trim());
  if (lines.length < 2) return []; // header + at least 1 row

  const header = lines[0].split(",").map((h) => h.trim().toLowerCase());
  const nameIdx = header.indexOf("name");
  const emailIdx = header.indexOf("email");

  if (nameIdx === -1 || emailIdx === -1) {
    throw new Error("CSV must have 'name' and 'email' columns");
  }

  const roleIdx = header.indexOf("role");
  const officeIdx = header.indexOf("office");

  const rows: CsvRow[] = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = lines[i].split(",").map((c) => c.trim());
    const name = cols[nameIdx] || "";
    const email = cols[emailIdx] || "";
    if (!name || !email) continue;

    rows.push({
      name,
      email: email.toLowerCase(),
      role: roleIdx !== -1 ? cols[roleIdx] : undefined,
      office: officeIdx !== -1 ? cols[officeIdx] : undefined,
    });
  }

  return rows;
}

/** POST /api/admin/users/import — bulk import users from CSV */
export async function POST(req: NextRequest) {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const formData = await req.formData();
  const file = formData.get("file") as File | null;

  if (!file) {
    return NextResponse.json({ error: "No file provided" }, { status: 400 });
  }

  if (!file.name.endsWith(".csv")) {
    return NextResponse.json({ error: "File must be a .csv" }, { status: 400 });
  }

  if (file.size > MAX_FILE_BYTES) {
    return NextResponse.json(
      { error: `File too large. Max ${MAX_FILE_BYTES / 1000} KB.` },
      { status: 413 }
    );
  }

  const text = await file.text();
  let rows: CsvRow[];

  try {
    rows = parseCsv(text);
  } catch (err) {
    const msg = err instanceof Error ? err.message : "Invalid CSV";
    return NextResponse.json({ error: msg }, { status: 400 });
  }

  if (rows.length === 0) {
    return NextResponse.json({ error: "CSV has no valid rows" }, { status: 400 });
  }

  if (rows.length > MAX_ROWS) {
    return NextResponse.json(
      { error: `Too many rows. Max ${MAX_ROWS} per import.` },
      { status: 413 }
    );
  }

  const supabase = createAdminClient();

  // Fetch existing emails
  const { data: existing } = await supabase.from("profiles").select("email");
  const existingEmails = new Set((existing ?? []).map((p) => p.email.toLowerCase()));

  let created = 0;
  let skipped = 0;
  const errors: { row: number; email: string; reason: string }[] = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    if (!EMAIL_RE.test(row.email)) {
      errors.push({ row: i + 2, email: row.email, reason: "Invalid email format" });
      continue;
    }

    if (!row.email.endsWith(`@${ALLOWED_DOMAIN}`)) {
      errors.push({ row: i + 2, email: row.email, reason: `Email must be @${ALLOWED_DOMAIN}` });
      continue;
    }

    if (existingEmails.has(row.email)) {
      skipped++;
      continue;
    }

    const role = row.role && VALID_ROLES.includes(row.role) ? row.role : "viewer";
    const office = row.office && VALID_OFFICES.includes(row.office) ? row.office : null;

    const { error } = await supabase.from("profiles").insert({
      name: row.name,
      email: row.email,
      role,
      office,
    });

    if (error) {
      errors.push({ row: i + 2, email: row.email, reason: error.message });
    } else {
      created++;
      existingEmails.add(row.email);
    }
  }

  await logAudit({
    userId: session.user.profileId,
    userEmail: session.user.email,
    action: "import_users_csv",
    resourceType: "profile",
    details: {
      filename: file.name,
      total_rows: rows.length,
      created,
      skipped,
      error_count: errors.length,
    },
  });

  return NextResponse.json({
    total: rows.length,
    created,
    skipped,
    errors,
  });
}
