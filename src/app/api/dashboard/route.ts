import { NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";

/** GET /api/dashboard — sales rep dashboard stats */
export async function GET() {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { role, profileId, office } = session.user;
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

  // Recent activity (last 10 quotes)
  const recentQuotes = quotes.slice(0, 10);
  // Recent customers (last 10)
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
    recentCustomers,
    monthlyRevenue,
  });
}
