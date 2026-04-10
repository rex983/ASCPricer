import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { logAudit } from "@/lib/audit";

/** GET /api/admin/sales-reps — list all sales reps with stats */
export async function GET() {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const supabase = createAdminClient();

  const { data: reps, error } = await supabase
    .from("asc_sales_reps")
    .select("*")
    .order("name");

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  // Fetch customer counts per rep
  const { data: customerCounts } = await supabase
    .from("asc_customers")
    .select("assigned_rep_id");

  // Fetch quote counts per rep (via created_by matching profile_id)
  const { data: quotes } = await supabase
    .from("asc_quotes")
    .select("created_by, status, total");

  // Build stats maps
  const customerCountMap: Record<string, number> = {};
  for (const c of customerCounts ?? []) {
    if (c.assigned_rep_id) {
      customerCountMap[c.assigned_rep_id] = (customerCountMap[c.assigned_rep_id] || 0) + 1;
    }
  }

  const quoteStatsMap: Record<string, { count: number; total: number }> = {};
  for (const q of quotes ?? []) {
    if (q.created_by) {
      if (!quoteStatsMap[q.created_by]) {
        quoteStatsMap[q.created_by] = { count: 0, total: 0 };
      }
      quoteStatsMap[q.created_by].count += 1;
      quoteStatsMap[q.created_by].total += q.total || 0;
    }
  }

  const enriched = (reps ?? []).map((r) => ({
    ...r,
    customer_count: customerCountMap[r.id] || 0,
    quote_count: r.profile_id ? quoteStatsMap[r.profile_id]?.count || 0 : 0,
    quote_total: r.profile_id ? quoteStatsMap[r.profile_id]?.total || 0 : 0,
  }));

  return NextResponse.json(enriched);
}

/** POST /api/admin/sales-reps — create a new sales rep */
export async function POST(req: NextRequest) {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const body = await req.json();
  const { name, email, phone, office, territory, commission_rate, profile_id, employee_role } = body as {
    name: string;
    email: string;
    phone?: string;
    office: string;
    territory?: string;
    commission_rate?: number;
    profile_id?: string;
    employee_role?: string;
  };

  if (!name?.trim()) {
    return NextResponse.json({ error: "Name is required" }, { status: 400 });
  }
  if (!email?.trim()) {
    return NextResponse.json({ error: "Email is required" }, { status: 400 });
  }
  if (!office || (office !== "Harbor" && office !== "Marion")) {
    return NextResponse.json({ error: "Office must be Harbor or Marion" }, { status: 400 });
  }

  const supabase = createAdminClient();

  const VALID_ROLES = ["sales_rep", "sales_manager", "bst"];

  const insertData: Record<string, unknown> = {
    name: name.trim(),
    email: email.trim(),
    phone: phone || null,
    office,
    territory: territory || null,
    commission_rate: commission_rate ?? 0,
    is_active: true,
    employee_role: employee_role && VALID_ROLES.includes(employee_role) ? employee_role : "sales_rep",
  };

  if (profile_id) {
    const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (UUID_RE.test(profile_id)) {
      insertData.profile_id = profile_id;
    }
  }

  const { data, error } = await supabase
    .from("asc_sales_reps")
    .insert(insertData)
    .select()
    .single();

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  await logAudit({
    userId: session.user.profileId,
    userEmail: session.user.email,
    action: "create_sales_rep",
    resourceType: "sales_rep",
    resourceId: data.id,
    details: { name: data.name, email: data.email, office },
  });

  return NextResponse.json(data, { status: 201 });
}
