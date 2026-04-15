import { NextRequest, NextResponse } from "next/server";
import { createAdminClient } from "@/lib/supabase/admin";
import { getImpersonationContext } from "@/lib/impersonation";

const ALLOWED_LIMITS = [10, 25, 50, 100] as const;

/** GET /api/dashboard — sales rep dashboard stats */
export async function GET(req: NextRequest) {
  const ctx = await getImpersonationContext();
  if (!ctx) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const requested = Number(req.nextUrl.searchParams.get("limit"));
  const recentLimit = (ALLOWED_LIMITS as readonly number[]).includes(requested)
    ? requested
    : 10;

  const { role, profileId, office } = ctx.effective;
  const supabase = createAdminClient();

  // Build customer query based on role
  let customerQuery = supabase
    .from("asc_customers")
    .select("id, name, assigned_rep_id, created_at, created_by");

  if (role === "sales_rep" || role === "viewer") {
    customerQuery = customerQuery.eq("created_by", profileId);
  } else if (role === "manager" && office) {
    customerQuery = customerQuery.eq("office", office);
  }

  // Build quote query based on role
  let quoteQuery = supabase
    .from("asc_quotes")
    .select("id, quote_number, status, total, created_at, created_by, customer_id");

  if (role === "sales_rep" || role === "viewer") {
    quoteQuery = quoteQuery.eq("created_by", profileId);
  } else if (role === "manager" && office) {
    quoteQuery = quoteQuery.eq("office", office);
  }

  const [customersResult, quotesResult] = await Promise.all([
    customerQuery.order("created_at", { ascending: false }),
    quoteQuery.order("created_at", { ascending: false }),
  ]);

  const customers = customersResult.data || [];
  const quotes = quotesResult.data || [];

  // Compute stats
  const totalCustomers = customers.length;
  const totalQuotes = quotes.length;
  const totalRevenue = quotes.reduce((sum, q) => sum + (q.total || 0), 0);

  const statusCounts: Record<string, number> = {};
  for (const q of quotes) {
    statusCounts[q.status] = (statusCounts[q.status] || 0) + 1;
  }

  // Recent activity — scope matches role: sales_rep = own, manager = office,
  // admin = all. Caller picks how many via ?limit=10|25|50|100.
  const recentQuotes = quotes.slice(0, recentLimit);
  const recentCustomers = customers.slice(0, 10);

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
    totalCustomers,
    totalQuotes,
    totalRevenue,
    statusCounts,
    recentQuotes,
    recentQuotesLimit: recentLimit,
    recentCustomers,
    monthlyRevenue,
  });
}
