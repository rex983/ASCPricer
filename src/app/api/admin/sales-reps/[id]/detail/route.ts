import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";

type Ctx = { params: Promise<{ id: string }> };

/** GET /api/admin/sales-reps/[id]/detail — full profile with quotes, customers, activity */
export async function GET(_req: NextRequest, ctx: Ctx) {
  const session = await auth();
  if (!session?.user || (session.user.role !== "admin" && session.user.role !== "manager")) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { id } = await ctx.params;
  const supabase = createAdminClient();

  // Fetch the rep
  const { data: rep, error: repError } = await supabase
    .from("asc_sales_reps")
    .select("*")
    .eq("id", id)
    .single();

  if (repError || !rep) {
    return NextResponse.json({ error: "Sales rep not found" }, { status: 404 });
  }

  // Fetch quotes, customers, and audit activity in parallel
  // Quotes are linked via created_by = rep.profile_id
  // Customers are linked via assigned_rep_id = rep.id
  // Audit entries are linked via user_id = rep.profile_id OR resource_id = rep.id
  const profileId = rep.profile_id;

  const [quotesResult, customersResult, activityResult] = await Promise.all([
    profileId
      ? supabase
          .from("asc_quotes")
          .select("id, quote_number, status, customer_name, total, created_at")
          .eq("created_by", profileId)
          .order("created_at", { ascending: false })
          .limit(50)
      : Promise.resolve({ data: [] }),
    Promise.resolve({ data: [] }),
    profileId
      ? supabase
          .from("asc_audit_log")
          .select("id, action, resource_type, resource_id, details, created_at")
          .eq("user_id", profileId)
          .order("created_at", { ascending: false })
          .limit(30)
      : Promise.resolve({ data: [] }),
  ]);

  const quotes = quotesResult.data ?? [];
  const customers = customersResult.data ?? [];
  const activity = activityResult.data ?? [];

  // Compute KPIs
  const totalQuotes = quotes.length;
  const totalRevenue = quotes.reduce((sum, q) => sum + (q.total || 0), 0);
  const acceptedQuotes = quotes.filter((q) => q.status === "accepted").length;
  const conversionRate = totalQuotes > 0 ? (acceptedQuotes / totalQuotes) * 100 : 0;

  const statusCounts: Record<string, number> = {};
  for (const q of quotes) {
    statusCounts[q.status] = (statusCounts[q.status] || 0) + 1;
  }

  // Monthly revenue (last 6 months)
  const monthlyRevenue: { month: string; revenue: number; count: number }[] = [];
  const now = new Date();
  for (let i = 5; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const label = d.toLocaleString("en-US", { month: "short", year: "2-digit" });
    const monthStart = d.toISOString();
    const monthEnd = new Date(d.getFullYear(), d.getMonth() + 1, 1).toISOString();
    const monthQuotes = quotes.filter(
      (q) => q.created_at >= monthStart && q.created_at < monthEnd
    );
    monthlyRevenue.push({
      month: label,
      revenue: monthQuotes.reduce((s, q) => s + (q.total || 0), 0),
      count: monthQuotes.length,
    });
  }

  return NextResponse.json({
    rep,
    quotes,
    customers,
    activity,
    kpis: {
      totalQuotes,
      totalRevenue,
      totalCustomers: customers.length,
      acceptedQuotes,
      conversionRate,
      statusCounts,
    },
    monthlyRevenue,
  });
}
