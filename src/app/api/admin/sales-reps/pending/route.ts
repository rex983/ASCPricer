import { NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";

/** GET /api/admin/sales-reps/pending — profiles not yet linked to a team member */
export async function GET() {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const supabase = createAdminClient();

  // Get all profile IDs that are already linked to a sales rep
  const { data: reps } = await supabase
    .from("asc_sales_reps")
    .select("profile_id")
    .not("profile_id", "is", null);

  const linkedIds = new Set((reps ?? []).map((r) => r.profile_id));

  // Also exclude the hardcoded admin account
  linkedIds.add("admin-001");
  linkedIds.add("dev-user-001");

  // Get all profiles
  const { data: profiles, error } = await supabase
    .from("profiles")
    .select("id, email, full_name, role, office, created_at")
    .order("created_at", { ascending: false });

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  const pending = (profiles ?? [])
    .filter((p) => !linkedIds.has(p.id))
    .map((p) => ({ ...p, name: p.full_name }));

  return NextResponse.json(pending);
}
