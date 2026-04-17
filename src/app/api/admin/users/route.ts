import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { logAudit } from "@/lib/audit";

const ALLOWED_ROLES = ["admin", "manager"];
const VALID_USER_ROLES = ["admin", "manager", "sales_rep", "bst"];
const VALID_OFFICES = ["Harbor", "Marion"];
/** GET /api/admin/users — list all profiles with sales rep data + stats */
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
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  const profiles = data ?? [];
  const profileIds = profiles.map((p) => p.id).filter(Boolean);

  // Fetch linked sales reps
  let repsMap: Record<string, {
    id: string;
    phone: string | null;
    is_active: boolean;
  }> = {};

  if (profileIds.length > 0) {
    const { data: reps } = await supabase
      .from("asc_sales_reps")
      .select("id, profile_id, phone, is_active")
      .in("profile_id", profileIds);

    for (const r of reps ?? []) {
      if (r.profile_id) {
        repsMap[r.profile_id] = {
          id: r.id,
          phone: r.phone,
          is_active: r.is_active,
        };
      }
    }
  }

  const customerCountMap: Record<string, number> = {};

  // Fetch quote stats per profile (quotes use created_by = profile_id)
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

  const enriched = profiles.map((p) => {
    const rep = repsMap[p.id] ?? null;
    return {
      ...p,
      // DB column is full_name; expose as `name` for the UI
      name: p.full_name ?? null,
      // Sales rep fields (null if no linked rep)
      rep_id: rep?.id ?? null,
      phone: rep?.phone ?? null,
      is_active: rep?.is_active ?? null,
      // Stats
      customer_count: rep ? (customerCountMap[rep.id] || 0) : 0,
      quote_count: quoteStatsMap[p.id]?.count || 0,
      quote_total: quoteStatsMap[p.id]?.total || 0,
    };
  });

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
    return NextResponse.json({ error: error.message }, { status: 500 });
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
