import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { logAudit } from "@/lib/audit";

const ALLOWED_ROLES = ["admin", "manager"];
const VALID_USER_ROLES = ["admin", "manager", "sales_rep", "viewer"];
const VALID_OFFICES = ["Harbor", "Marion"];

/** GET /api/admin/users — list all profiles */
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

  // Check which profiles are linked to a sales rep
  const profileIds = (data ?? []).map((p) => p.id).filter(Boolean);
  let linkedRepIds: Set<string> = new Set();
  if (profileIds.length > 0) {
    const { data: reps } = await supabase
      .from("asc_sales_reps")
      .select("profile_id")
      .in("profile_id", profileIds);
    linkedRepIds = new Set((reps ?? []).map((r) => r.profile_id).filter(Boolean));
  }

  const enriched = (data ?? []).map((p) => ({
    ...p,
    has_sales_rep: linkedRepIds.has(p.id),
  }));

  return NextResponse.json({ users: enriched, total: count ?? 0 });
}

/** POST /api/admin/users — create a new profile (invite user) */
export async function POST(req: NextRequest) {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const body = await req.json();
  const { name, email, role, office } = body as {
    name: string;
    email: string;
    role: string;
    office?: string;
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

  const { data, error } = await supabase
    .from("profiles")
    .insert({
      name: name.trim(),
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
    resourceId: data.id,
    details: { name: data.name, email: data.email, role: data.role, office: data.office },
  });

  return NextResponse.json(data, { status: 201 });
}
