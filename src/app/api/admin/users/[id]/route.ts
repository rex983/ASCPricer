import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { logAudit } from "@/lib/audit";

type Ctx = { params: Promise<{ id: string }> };

const ALLOWED_ROLES = ["admin", "manager"];
const VALID_USER_ROLES = ["admin", "manager", "sales_rep", "viewer"];
const VALID_OFFICES = ["Harbor", "Marion"];

/** PATCH /api/admin/users/[id] — update a user profile */
export async function PATCH(req: NextRequest, ctx: Ctx) {
  const session = await auth();
  if (!session?.user || !ALLOWED_ROLES.includes(session.user.role)) {
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
    updates.email = body.email.trim().toLowerCase();
  }

  if (body.role !== undefined) {
    // Only admins can change roles
    if (session.user.role !== "admin") {
      return NextResponse.json({ error: "Only admins can change user roles" }, { status: 403 });
    }
    if (!VALID_USER_ROLES.includes(body.role)) {
      return NextResponse.json(
        { error: `Role must be one of: ${VALID_USER_ROLES.join(", ")}` },
        { status: 400 }
      );
    }
    updates.role = body.role;
  }

  if (body.office !== undefined) {
    if (body.office && !VALID_OFFICES.includes(body.office)) {
      return NextResponse.json({ error: "Office must be Harbor or Marion" }, { status: 400 });
    }
    updates.office = body.office || null;
  }

  if (Object.keys(updates).length === 0) {
    return NextResponse.json({ error: "No updates provided" }, { status: 400 });
  }

  const supabase = createAdminClient();

  // Prevent managers from editing users outside their office
  if (session.user.role === "manager" && session.user.office) {
    const { data: target } = await supabase
      .from("profiles")
      .select("office")
      .eq("id", id)
      .single();

    if (target && target.office !== session.user.office) {
      return NextResponse.json({ error: "Cannot edit users from another office" }, { status: 403 });
    }
  }

  // Prevent editing the hardcoded admin
  const { data: targetUser } = await supabase
    .from("profiles")
    .select("email")
    .eq("id", id)
    .single();

  if (targetUser?.email === "rex@bigbuildingsdirect.com" && updates.role && updates.role !== "admin") {
    return NextResponse.json({ error: "Cannot change the primary admin's role" }, { status: 403 });
  }

  // Check email uniqueness if changing email
  if (updates.email) {
    const { data: dup } = await supabase
      .from("profiles")
      .select("id")
      .eq("email", updates.email)
      .neq("id", id)
      .maybeSingle();

    if (dup) {
      return NextResponse.json({ error: "A user with this email already exists" }, { status: 409 });
    }
  }

  const { data, error } = await supabase
    .from("profiles")
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
    action: "update_user",
    resourceType: "profile",
    resourceId: id,
    details: updates,
  });

  return NextResponse.json(data);
}

/** DELETE /api/admin/users/[id] — permanently delete a user profile */
export async function DELETE(_req: NextRequest, ctx: Ctx) {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { id } = await ctx.params;
  const supabase = createAdminClient();

  // Prevent deleting yourself
  if (id === session.user.profileId) {
    return NextResponse.json({ error: "Cannot delete your own account" }, { status: 400 });
  }

  // Prevent deleting the primary admin
  const { data: target } = await supabase
    .from("profiles")
    .select("email, name")
    .eq("id", id)
    .single();

  if (!target) {
    return NextResponse.json({ error: "User not found" }, { status: 404 });
  }

  if (target.email === "rex@bigbuildingsdirect.com") {
    return NextResponse.json({ error: "Cannot delete the primary admin account" }, { status: 403 });
  }

  // Check for linked sales rep
  const { count: repCount } = await supabase
    .from("asc_sales_reps")
    .select("id", { count: "exact", head: true })
    .eq("profile_id", id);

  if (repCount && repCount > 0) {
    return NextResponse.json(
      { error: "Cannot delete: this user is linked to a sales rep. Remove the link first." },
      { status: 409 }
    );
  }

  // Check for quotes created by this user
  const { count: quoteCount } = await supabase
    .from("asc_quotes")
    .select("id", { count: "exact", head: true })
    .eq("created_by", id);

  if (quoteCount && quoteCount > 0) {
    return NextResponse.json(
      { error: `Cannot delete: this user has ${quoteCount} quote(s). Reassign or delete them first.` },
      { status: 409 }
    );
  }

  const { error } = await supabase.from("profiles").delete().eq("id", id);

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  await logAudit({
    userId: session.user.profileId,
    userEmail: session.user.email,
    action: "delete_user",
    resourceType: "profile",
    resourceId: id,
    details: { name: target.name, email: target.email },
  });

  return NextResponse.json({ success: true });
}
