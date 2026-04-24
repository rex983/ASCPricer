import { NextRequest, NextResponse } from "next/server";
import { createAdminClient } from "@/lib/supabase/admin";
import { getImpersonationContext } from "@/lib/impersonation";

const ALLOWED_LIMITS = [10, 25, 50, 100] as const;
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

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

  // Date range — default to current month
  const now = new Date();
  const defaultStart = new Date(now.getFullYear(), now.getMonth(), 1)
    .toISOString()
    .slice(0, 10);
  const defaultEnd = now.toISOString().slice(0, 10);

  const startParam = req.nextUrl.searchParams.get("startDate");
  const endParam = req.nextUrl.searchParams.get("endDate");
  const startDate = startParam && DATE_RE.test(startParam) ? startParam : defaultStart;
  const endDate = endParam && DATE_RE.test(endParam) ? endParam : defaultEnd;

  // Convert to ISO timestamps — endDate is inclusive (through end of day)
  const startISO = `${startDate}T00:00:00.000Z`;
  const endISO = `${endDate}T23:59:59.999Z`;

  const { role, profileId, office } = ctx.effective;
  const supabase = createAdminClient();

  // Build customer query based on role + date range
  let customerQuery = supabase
    .from("asc_customers")
    .select("id, name, created_at, created_by")
    .gte("created_at", startISO)
    .lte("created_at", endISO);

  if (role === "sales_rep" || role === "bst") {
    customerQuery = customerQuery.eq("created_by", profileId);
  } else if (role === "manager" && office) {
    customerQuery = customerQuery.eq("office", office);
  }

  // Build quote query based on role + date range
  let quoteQuery = supabase
    .from("asc_quotes")
    .select("id, quote_number, status, total, created_at, created_by, customer_id")
    .gte("created_at", startISO)
    .lte("created_at", endISO);

  if (role === "sales_rep" || role === "bst") {
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

  const recentQuotes = quotes.slice(0, recentLimit);
  const recentCustomers = customers.slice(0, 10);

  // Monthly revenue — build buckets spanning the selected date range
  const rangeStart = new Date(startDate + "T00:00:00Z");
  const rangeEnd = new Date(endDate + "T23:59:59Z");
  const monthlyRevenue: { month: string; revenue: number; count: number }[] = [];
  const cursor = new Date(rangeStart.getFullYear(), rangeStart.getMonth(), 1);

  while (cursor <= rangeEnd) {
    const label = cursor.toLocaleString("en-US", { month: "short", year: "2-digit" });
    const monthStart = cursor.toISOString();
    const monthEnd = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1).toISOString();
    const monthQuotes = quotes.filter(
      (q) => q.created_at >= monthStart && q.created_at < monthEnd
    );
    monthlyRevenue.push({
      month: label,
      revenue: monthQuotes.reduce((s, q) => s + (q.total || 0), 0),
      count: monthQuotes.length,
    });
    cursor.setMonth(cursor.getMonth() + 1);
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
    startDate,
    endDate,
  });
}
