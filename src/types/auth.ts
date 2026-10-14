import type { DefaultSession } from "next-auth";

/**
 * Roles come from the launcher's `profiles.role` (launcher_roles). The four
 * named here are the ones this app grants something to; any other launcher
 * role (junior_manager, team_lead, rnd, custom roles, ...) is valid and gets
 * the least-privileged behavior (own records only, no admin pages).
 */
export type KnownRole = "admin" | "senior_manager" | "sales_rep" | "bst";
export type UserRole = KnownRole | (string & {});
/** Launcher office name (launcher_offices) — admins can add more at any time. */
export type Office = string;

/** Launcher admin "View as" (read-only impersonation) — set from the SSO act claim. */
export interface ViewAs {
  actorProfileId: string;
  actorEmail: string;
  actorName: string;
  startedAt: number; // epoch ms
}

declare module "next-auth" {
  interface User {
    /** Launcher SSO actor claim (admin "View as"); only set by launcher-sso. */
    act?: { sub: string; profile_id: string; name: string } | null;
  }

  interface Session {
    user: {
      /** Real launcher role — display only. */
      role: UserRole;
      /** launcher_permission_role(role): use for EVERY in-app permission check. */
      permRole: UserRole;
      profileId: string;
      office?: Office;
      department?: string | null;
      isIt?: boolean;
    } & DefaultSession["user"];
    /** Non-null while a launcher admin is viewing as this user (read-only). */
    viewAs: ViewAs | null;
  }

  interface JWT {
    role?: UserRole;
    permRole?: UserRole;
    accessOkAt?: number;
    profileId?: string;
    office?: Office;
    department?: string | null;
    isIt?: boolean;
    viewAs?: ViewAs | null;
  }
}
