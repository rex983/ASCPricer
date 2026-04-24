import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { logAudit } from "@/lib/audit";

const ALLOWED_ROLES = ["admin", "manager"];
const VALID_USER_ROLES = ["admin", "manager", "sales_rep", "bst"];
const VALID_OFFICES = ["Harbor", "Marion"];
/** GET /api/admin/users — list all profiles + stats */
export async function GET(req: NextRequest) {
  const session = await auth();
  if (!session?.user || !ALLOWED_ROLES.includes(session.user.role)) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const supabase = createAdminClient();
  const { searchParams } = req.nextUrl;
  const limit = Math.min(Number(searchParams.get("limit")) || 100, 200);
  const offset = Number(searchParams.get("offset")) || 0;

  let query = supabase
    .from("profiles")
    .select("*", { count: "exact" })
    .order("created_at", { ascending: false })
    .range(offset, offset + limit - 1);

  // Managers can only see their own office's users
  if (session.user.role === "manager" && session.user.office) {
    query = query.eq("office", session.user.office);
  }

  const { data, error, count } = await query;

  if (error) {
    console.error("users GET error:", error);
    return NextResponse.json({ error: "Database operation failed" }, { status: 500 });
  }

  const profiles = data ?? [];
  const profileIds = profiles.map((p) => p.id).filter(Boolean);

  // Fetch quote stats per profile
  const quoteStatsMap: Record<string, { count: number; total: number }> = {};
  if (profileIds.length > 0) {
    const { data: quotes } = await supabase
      .from("asc_quotes")
      .select("created_by, total");

    for (const q of quotes ?? []) {
      if (q.created_by) {
        if (!quoteStatsMap[q.created_by]) {
          quoteStatsMap[q.created_by] = { count: 0, total: 0 };
        }
        quoteStatsMap[q.created_by].count += 1;
        quoteStatsMap[q.created_by].total += q.total || 0;
      }
    }
  }

  const enriched = profiles.map((p) => ({
    ...p,
    name: p.full_name ?? null,
    quote_count: quoteStatsMap[p.id]?.count || 0,
    quote_total: quoteStatsMap[p.id]?.total || 0,
  }));

  return NextResponse.json({ users: enriched, total: count ?? 0 });
}

/** POST /api/admin/users — create a new profile + optionally a linked sales rep */
export async function POST(req: NextRequest) {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const body = await req.json();
  const { name, email, role, office, phone } = body as {
    name: string;
    email: string;
    role: string;
    office?: string;
    phone?: string;
  };

  if (!name?.trim()) {
    return NextResponse.json({ error: "Name is required" }, { status: 400 });
  }
  if (!email?.trim()) {
    return NextResponse.json({ error: "Email is required" }, { status: 400 });
  }
  if (!role || !VALID_USER_ROLES.includes(role)) {
    return NextResponse.json(
      { error: `Role must be one of: ${VALID_USER_ROLES.join(", ")}` },
      { status: 400 }
    );
  }
  if (office && !VALID_OFFICES.includes(office)) {
    return NextResponse.json({ error: "Office must be Harbor or Marion" }, { status: 400 });
  }

  const supabase = createAdminClient();

  // Check for duplicate email
  const { data: existing } = await supabase
    .from("profiles")
    .select("id")
    .eq("email", email.trim().toLowerCase())
    .maybeSingle();

  if (existing) {
    return NextResponse.json({ error: "A user with this email already exists" }, { status: 409 });
  }

  const { data: profile, error } = await supabase
    .from("profiles")
    .insert({
      full_name: name.trim(),
      email: email.trim().toLowerCase(),
      role,
      office: office || null,
    })
    .select()
    .single();

  if (error) {
    console.error("users POST error:", error);
    return NextResponse.json({ error: "Database operation failed" }, { status: 500 });
  }

  await logAudit({
    userId: session.user.profileId,
    userEmail: session.user.email,
    action: "create_user",
    resourceType: "profile",
    resourceId: profile.id,
    details: { name: profile.full_name, email: profile.email, role: profile.role, office: profile.office },
  });

  return NextResponse.json({ ...profile, name: profile.full_name }, { status: 201 });
}
