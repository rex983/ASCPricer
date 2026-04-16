import { cookies } from "next/headers";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import type { UserRole, Office } from "@/types/auth";

export const IMPERSONATION_COOKIE = "view_as_profile_id";
export const IMPERSONATION_MAX_AGE = 8 * 60 * 60;

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
const IMPERSONATABLE_ROLES: UserRole[] = ["sales_rep", "bst"];

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
  const targetId = jar.get(IMPERSONATION_COOKIE)?.value;

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
