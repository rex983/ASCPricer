import { cookies } from "next/headers";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { createHmac } from "crypto";
import type { UserRole, Office } from "@/types/auth";

export const IMPERSONATION_COOKIE = "view_as_profile_id";
export const IMPERSONATION_MAX_AGE = 8 * 60 * 60;

/**
 * Sign an impersonation cookie value so it's bound to the real user's session.
 * Format: "targetId.hmac" where hmac = HMAC(realProfileId:targetId, secret)
 */
function getHmacSecret(): string {
  return process.env.AUTH_SECRET || process.env.NEXTAUTH_SECRET || "fallback-dev-secret";
}

export function signCookieValue(realProfileId: string, targetId: string): string {
  const mac = createHmac("sha256", getHmacSecret())
    .update(`${realProfileId}:${targetId}`)
    .digest("hex")
    .slice(0, 16);
  return `${targetId}.${mac}`;
}

export function verifyCookieValue(cookieValue: string, realProfileId: string): string | null {
  const dotIdx = cookieValue.indexOf(".");
  if (dotIdx === -1) return null; // unsigned legacy cookie — reject
  const targetId = cookieValue.slice(0, dotIdx);
  const providedMac = cookieValue.slice(dotIdx + 1);
  const expectedMac = createHmac("sha256", getHmacSecret())
    .update(`${realProfileId}:${targetId}`)
    .digest("hex")
    .slice(0, 16);
  if (providedMac !== expectedMac) return null;
  return targetId;
}

export interface EffectiveUser {
  role: UserRole;
  profileId: string;
  office?: Office;
  email?: string | null;
  name?: string | null;
}

export interface ImpersonationTarget {
  profileId: string;
  name: string | null;
  email: string;
  office: Office | null;
  role: UserRole;
}

export interface ImpersonationContext {
  real: EffectiveUser;
  effective: EffectiveUser;
  isImpersonating: boolean;
  target: ImpersonationTarget | null;
  canImpersonate: boolean;
}

// Only these roles can be impersonation targets — never admin/manager.
const IMPERSONATABLE_ROLES: UserRole[] = ["sales_rep", "bst", "manager"];

/**
 * Resolve the real session user and the effective user (after any valid
 * impersonation cookie is applied). Re-validates authorization on every call
 * so a stale cookie can never elevate privileges.
 */
export async function getImpersonationContext(): Promise<ImpersonationContext | null> {
  const session = await auth();
  if (!session?.user) return null;

  const real: EffectiveUser = {
    role: session.user.role,
    profileId: session.user.profileId,
    office: session.user.office,
    email: session.user.email ?? null,
    name: session.user.name ?? null,
  };

  const canImpersonate = real.role === "admin" || real.role === "manager";

  const jar = await cookies();
  const rawCookie = jar.get(IMPERSONATION_COOKIE)?.value;
  const targetId = rawCookie ? verifyCookieValue(rawCookie, real.profileId) : null;

  if (!targetId || !canImpersonate) {
    return { real, effective: real, isImpersonating: false, target: null, canImpersonate };
  }

  const supabase = createAdminClient();
  const { data: target } = await supabase
    .from("profiles")
    .select("id, full_name, email, role, office")
    .eq("id", targetId)
    .single();

  const targetRole = target?.role as UserRole | undefined;
  const targetOffice = (target?.office as Office | null) ?? null;

  const isValid =
    !!target &&
    !!targetRole &&
    IMPERSONATABLE_ROLES.includes(targetRole) &&
    // Managers can only view reps in their own office.
    (real.role === "admin" || (real.office && targetOffice === real.office));

  if (!isValid || !target || !targetRole) {
    return { real, effective: real, isImpersonating: false, target: null, canImpersonate };
  }

  const effective: EffectiveUser = {
    role: targetRole,
    profileId: target.id,
    office: targetOffice ?? undefined,
    email: target.email,
    name: target.full_name,
  };

  return {
    real,
    effective,
    isImpersonating: true,
    target: {
      profileId: target.id,
      name: target.full_name,
      email: target.email,
      office: targetOffice,
      role: targetRole,
    },
    canImpersonate,
  };
}

/** Convenience: return the effective user for role-based query filtering. */
export async function getEffectiveUser(): Promise<EffectiveUser | null> {
  const ctx = await getImpersonationContext();
  return ctx?.effective ?? null;
}
