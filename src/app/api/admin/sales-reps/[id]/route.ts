import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { logAudit } from "@/lib/audit";

type Ctx = { params: Promise<{ id: string }> };

/** PATCH /api/admin/sales-reps/[id] — update sales rep fields */
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

  if (body.email !== undefined) {
    if (!body.email?.trim()) {
      return NextResponse.json({ error: "Email cannot be empty" }, { status: 400 });
    }
    updates.email = body.email.trim();
  }

  if (body.phone !== undefined) updates.phone = body.phone || null;
  if (body.office !== undefined) {
    if (body.office !== "Harbor" && body.office !== "Marion") {
      return NextResponse.json({ error: "Office must be Harbor or Marion" }, { status: 400 });
    }
    updates.office = body.office;
  }
  if (body.is_active !== undefined) updates.is_active = Boolean(body.is_active);
  if (body.profile_id !== undefined) {
    const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    updates.profile_id = body.profile_id && UUID_RE.test(body.profile_id) ? body.profile_id : null;
  }

  if (Object.keys(updates).length === 0) {
    return NextResponse.json({ error: "No updates provided" }, { status: 400 });
  }

  const supabase = createAdminClient();

  const { data, error } = await supabase
    .from("asc_sales_reps")
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
    action: body.is_active !== undefined ? "toggle_sales_rep" : "update_sales_rep",
    resourceType: "sales_rep",
    resourceId: id,
    details: updates,
  });

  return NextResponse.json(data);
}

/** DELETE /api/admin/sales-reps/[id] — permanently delete a sales rep */
export async function DELETE(_req: NextRequest, ctx: Ctx) {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { id } = await ctx.params;
  const supabase = createAdminClient();

  const { error } = await supabase
    .from("asc_sales_reps")
    .delete()
    .eq("id", id);

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  await logAudit({
    userId: session.user.profileId,
    userEmail: session.user.email,
    action: "delete_sales_rep",
    resourceType: "sales_rep",
    resourceId: id,
  });

  return NextResponse.json({ success: true });
}
