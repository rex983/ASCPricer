import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { logAudit } from "@/lib/audit";

const VALID_STATES = new Set([
  "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN",
  "IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV",
  "NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN",
  "TX","UT","VT","VA","WA","WV","WI","WY",
]);

type Ctx = { params: Promise<{ id: string }> };

/** PATCH /api/admin/regions/[id] — update region fields */
export async function PATCH(req: NextRequest, ctx: Ctx) {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { id } = await ctx.params;
  const body = await req.json();
  const updates: Record<string, unknown> = {};

  if (body.name !== undefined) {
    if (!body.name?.trim()) {
      return NextResponse.json({ error: "Name cannot be empty" }, { status: 400 });
    }
    updates.name = body.name.trim();
  }

  if (body.states !== undefined) {
    if (!Array.isArray(body.states) || body.states.length === 0) {
      return NextResponse.json({ error: "At least one state is required" }, { status: 400 });
    }
    const invalid = body.states.filter((s: string) => !VALID_STATES.has(s));
    if (invalid.length > 0) {
      return NextResponse.json({ error: `Invalid state codes: ${invalid.join(", ")}` }, { status: 400 });
    }
    updates.states = body.states;
  }

  if (body.is_active !== undefined) {
    updates.is_active = Boolean(body.is_active);
  }

  if (Object.keys(updates).length === 0) {
    return NextResponse.json({ error: "No updates provided" }, { status: 400 });
  }

  const supabase = createAdminClient();

  const { data, error } = await supabase
    .from("asc_regions")
    .update(updates)
    .eq("id", id)
    .select()
    .single();

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  await logAudit({
    userId: session.user.profileId,
    userEmail: session.user.email,
    action: body.is_active !== undefined ? "toggle_region" : "update_region",
    resourceType: "region",
    resourceId: id,
    details: updates,
  });

  return NextResponse.json(data);
}

/** DELETE /api/admin/regions/[id] — permanently delete a region */
export async function DELETE(_req: NextRequest, ctx: Ctx) {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { id } = await ctx.params;
  const supabase = createAdminClient();

  // Check for quotes before deleting
  const { count } = await supabase
    .from("asc_quotes")
    .select("id", { count: "exact", head: true })
    .eq("region_id", id);

  if (count && count > 0) {
    return NextResponse.json(
      { error: `Cannot delete: ${count} quote(s) reference this region. Deactivate it instead.` },
      { status: 409 }
    );
  }

  const { error } = await supabase
    .from("asc_regions")
    .delete()
    .eq("id", id);

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  await logAudit({
    userId: session.user.profileId,
    userEmail: session.user.email,
    action: "delete_region",
    resourceType: "region",
    resourceId: id,
  });

  return NextResponse.json({ success: true });
}
