import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { logAudit } from "@/lib/audit";

type Ctx = { params: Promise<{ id: string }> };

const ALLOWED_ROLES = ["admin", "manager"];
const VALID_USER_ROLES = ["admin", "manager", "sales_rep", "viewer"];
const VALID_OFFICES = ["Harbor", "Marion"];
const VALID_EMPLOYEE_ROLES = ["sales_rep", "sales_manager", "bst"];

/** PATCH /api/admin/users/[id] — update profile + linked sales rep */
export async function PATCH(req: NextRequest, ctx: Ctx) {
  const session = await auth();
  if (!session?.user || !ALLOWED_ROLES.includes(session.user.role)) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { id } = await ctx.params;
  const body = await req.json();
  const profileUpdates: Record<string, unknown> = {};
  const repUpdates: Record<string, unknown> = {};

  // Profile fields
  if (body.name !== undefined) {
    if (!body.name?.trim()) {
      return NextResponse.json({ error: "Name cannot be empty" }, { status: 400 });
    }
    profileUpdates.full_name = body.name.trim();
    repUpdates.name = body.name.trim();
  }

  if (body.email !== undefined) {
    if (!body.email?.trim()) {
      return NextResponse.json({ error: "Email cannot be empty" }, { status: 400 });
    }
    profileUpdates.email = body.email.trim().toLowerCase();
    repUpdates.email = body.email.trim().toLowerCase();
  }

  if (body.role !== undefined) {
    if (session.user.role !== "admin") {
      return NextResponse.json({ error: "Only admins can change user roles" }, { status: 403 });
    }
    if (!VALID_USER_ROLES.includes(body.role)) {
      return NextResponse.json(
        { error: `Role must be one of: ${VALID_USER_ROLES.join(", ")}` },
        { status: 400 }
      );
    }
    profileUpdates.role = body.role;
  }

  if (body.office !== undefined) {
    if (body.office && !VALID_OFFICES.includes(body.office)) {
      return NextResponse.json({ error: "Office must be Harbor or Marion" }, { status: 400 });
    }
    profileUpdates.office = body.office || null;
    repUpdates.office = body.office || null;
  }

  // Sales rep fields
  if (body.phone !== undefined) repUpdates.phone = body.phone || null;
  if (body.territory !== undefined) repUpdates.territory = body.territory || null;
  if (body.commission_rate !== undefined) repUpdates.commission_rate = body.commission_rate;
  if (body.is_active !== undefined) repUpdates.is_active = Boolean(body.is_active);
  if (body.employee_role !== undefined) {
    if (body.employee_role && !VALID_EMPLOYEE_ROLES.includes(body.employee_role)) {
      return NextResponse.json(
        { error: `Employee role must be one of: ${VALID_EMPLOYEE_ROLES.join(", ")}` },
        { status: 400 }
      );
    }
    repUpdates.employee_role = body.employee_role;
  }

  const hasProfileUpdates = Object.keys(profileUpdates).length > 0;
  const hasRepUpdates = Object.keys(repUpdates).length > 0;

  if (!hasProfileUpdates && !hasRepUpdates) {
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

  // Prevent demoting the primary admin
  const { data: targetUser } = await supabase
    .from("profiles")
    .select("email, full_name")
    .eq("id", id)
    .single();

  if (targetUser?.email === "rex@bigbuildingsdirect.com" && profileUpdates.role && profileUpdates.role !== "admin") {
    return NextResponse.json({ error: "Cannot change the primary admin's role" }, { status: 403 });
  }

  // Check email uniqueness if changing email
  if (profileUpdates.email) {
    const { data: dup } = await supabase
      .from("profiles")
      .select("id")
      .eq("email", profileUpdates.email)
      .neq("id", id)
      .maybeSingle();

    if (dup) {
      return NextResponse.json({ error: "A user with this email already exists" }, { status: 409 });
    }
  }

  // Update profile
  let profile = null;
  if (hasProfileUpdates) {
    const { data, error } = await supabase
      .from("profiles")
      .update(profileUpdates)
      .eq("id", id)
      .select()
      .single();

    if (error) {
      return NextResponse.json({ error: error.message }, { status: 500 });
    }
    profile = data;
  }

  // Update or create linked sales rep
  if (hasRepUpdates) {
    const { data: existingRep } = await supabase
      .from("asc_sales_reps")
      .select("id")
      .eq("profile_id", id)
      .maybeSingle();

    if (existingRep) {
      // Update existing rep
      await supabase
        .from("asc_sales_reps")
        .update(repUpdates)
        .eq("profile_id", id);
    } else if (body.employee_role && VALID_EMPLOYEE_ROLES.includes(body.employee_role)) {
      // Create new rep — need full required fields
      const currentProfile = profile || targetUser;
      await supabase.from("asc_sales_reps").insert({
        profile_id: id,
        name: (repUpdates.name as string) || currentProfile?.full_name || "",
        email: (repUpdates.email as string) || currentProfile?.email || "",
        phone: repUpdates.phone ?? null,
        office: (repUpdates.office as string) || "Harbor",
        territory: repUpdates.territory ?? null,
        commission_rate: (repUpdates.commission_rate as number) ?? 0,
        employee_role: body.employee_role,
        is_active: true,
      });
    }
  }

  await logAudit({
    userId: session.user.profileId,
    userEmail: session.user.email,
    action: body.is_active !== undefined ? "toggle_user" : "update_user",
    resourceType: "profile",
    resourceId: id,
    details: { ...profileUpdates, ...repUpdates },
  });

  // Re-fetch the profile for the response
  if (!profile) {
    const { data } = await supabase.from("profiles").select("*").eq("id", id).single();
    profile = data;
  }

  return NextResponse.json(profile ? { ...profile, name: profile.full_name } : profile);
}

/** DELETE /api/admin/users/[id] — delete profile + linked sales rep */
export async function DELETE(_req: NextRequest, ctx: Ctx) {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { id } = await ctx.params;
  const supabase = createAdminClient();

  if (id === session.user.profileId) {
    return NextResponse.json({ error: "Cannot delete your own account" }, { status: 400 });
  }

  const { data: target } = await supabase
    .from("profiles")
    .select("email, full_name")
    .eq("id", id)
    .single();

  if (!target) {
    return NextResponse.json({ error: "User not found" }, { status: 404 });
  }

  if (target.email === "rex@bigbuildingsdirect.com") {
    return NextResponse.json({ error: "Cannot delete the primary admin account" }, { status: 403 });
  }

  // Check for assigned customers (via sales rep)
  const { data: rep } = await supabase
    .from("asc_sales_reps")
    .select("id")
    .eq("profile_id", id)
    .maybeSingle();

  if (rep) {
    const { count: custCount } = await supabase
      .from("asc_customers")
      .select("id", { count: "exact", head: true })
      .eq("assigned_rep_id", rep.id);

    if (custCount && custCount > 0) {
      return NextResponse.json(
        { error: `Cannot delete: ${custCount} customer(s) assigned. Reassign them first.` },
        { status: 409 }
      );
    }
  }

  // Check for quotes
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

  // Delete linked sales rep first (if exists)
  if (rep) {
    await supabase.from("asc_sales_reps").delete().eq("id", rep.id);
  }

  // Delete profile
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
    details: { name: target.full_name, email: target.email },
  });

  return NextResponse.json({ success: true });
}
