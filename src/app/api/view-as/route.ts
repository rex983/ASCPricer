import { NextRequest, NextResponse } from "next/server";
import { cookies } from "next/headers";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { logAudit } from "@/lib/audit";
import {
  IMPERSONATION_COOKIE,
  IMPERSONATION_MAX_AGE,
  getImpersonationContext,
} from "@/lib/impersonation";

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const IMPERSONATABLE_ROLES = ["sales_rep", "bst", "manager"];

/** GET /api/view-as — current impersonation status for the UI banner. */
export async function GET() {
  const ctx = await getImpersonationContext();
  if (!ctx) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }
  return NextResponse.json({
    isImpersonating: ctx.isImpersonating,
    target: ctx.target,
    canImpersonate: ctx.canImpersonate,
  });
}

/** POST /api/view-as — start viewing as a sales rep. */
export async function POST(req: NextRequest) {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { role: realRole, office: realOffice, profileId: realId, email: realEmail } =
    session.user;

  if (realRole !== "admin" && realRole !== "manager") {
    return NextResponse.json({ error: "Forbidden" }, { status: 403 });
  }

  const body = await req.json().catch(() => ({}));
  const targetProfileId = String(body?.targetProfileId || "");
  if (!UUID_RE.test(targetProfileId)) {
    return NextResponse.json({ error: "Invalid target" }, { status: 400 });
  }

  if (targetProfileId === realId) {
    return NextResponse.json(
      { error: "You cannot view as yourself" },
      { status: 400 }
    );
  }

  const supabase = createAdminClient();
  const { data: target } = await supabase
    .from("profiles")
    .select("id, full_name, email, role, office")
    .eq("id", targetProfileId)
    .single();

  if (!target) {
    return NextResponse.json({ error: "Target not found" }, { status: 404 });
  }

  if (!IMPERSONATABLE_ROLES.includes(target.role)) {
    return NextResponse.json(
      { error: "You can only view as sales reps" },
      { status: 400 }
    );
  }

  if (realRole === "manager") {
    if (!realOffice || target.office !== realOffice) {
      return NextResponse.json(
        { error: "You can only view reps in your own office" },
        { status: 403 }
      );
    }
  }

  const jar = await cookies();
  jar.set(IMPERSONATION_COOKIE, target.id, {
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
    path: "/",
    maxAge: IMPERSONATION_MAX_AGE,
  });

  await logAudit({
    userId: realId,
    userEmail: realEmail,
    action: "view_as_start",
    resourceType: "profile",
    resourceId: target.id,
    details: {
      targetEmail: target.email,
      targetRole: target.role,
      targetOffice: target.office,
    },
  });

  return NextResponse.json({
    ok: true,
    target: {
      profileId: target.id,
      name: target.full_name,
      email: target.email,
      office: target.office,
      role: target.role,
    },
  });
}

/** DELETE /api/view-as — stop impersonating. */
export async function DELETE() {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const jar = await cookies();
  const prev = jar.get(IMPERSONATION_COOKIE)?.value;
  jar.delete(IMPERSONATION_COOKIE);

  if (prev) {
    await logAudit({
      userId: session.user.profileId,
      userEmail: session.user.email,
      action: "view_as_stop",
      resourceType: "profile",
      resourceId: prev,
    });
  }

  return NextResponse.json({ ok: true });
}
