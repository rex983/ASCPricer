import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { LAUNCHER_URL } from "@/lib/launcher-access";

// Read-only. Users, roles and offices are managed only in the BBD Launcher
// (Admin -> Users); the old create / edit / delete / CSV-import endpoints
// that wrote the shared profiles table from here were removed.
const ALLOWED_ROLES = ["admin", "senior_manager"];
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
  if (session.user.role === "senior_manager" && session.user.office) {
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

  return NextResponse.json({
    users: enriched,
    total: count ?? 0,
    launcherUrl: LAUNCHER_URL,
  });
}
