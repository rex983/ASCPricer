import type { SupabaseClient } from "@supabase/supabase-js";
import { createAdminClient } from "@/lib/supabase/admin";

// Must match launcher_sso_configs.jwt_audience / launcher_apps for this app.
export const LAUNCHER_AUDIENCE = "asc-pricer";

// Display name used in access-denied messages.
export const APP_NAME = "ASC Pricer";

// Where users and access are managed. Used for "Managed in the BBD Launcher"
// links and the view-as Exit button.
export const LAUNCHER_URL = (
  process.env.LAUNCHER_URL || "https://bbd-launcher.vercel.app"
).replace(/\/+$/, "");

// Error codes put on /login?error=... by the signIn callback.
export const NO_LAUNCHER_ACCESS = "NoLauncherAccess";
export const LAUNCHER_CHECK_FAILED = "LauncherCheckFailed";

// A session whose access couldn't be re-confirmed (RPC missing / transient
// errors) survives at most this long after its last successful check.
export const ACCESS_GRACE_MS = 5 * 60 * 1000;

/**
 * Launcher access is fail-CLOSED:
 *   true    — launcher_can_open_app returned exactly true
 *   false   — returned false or NULL (app not registered) => deny
 *   "error" — RPC missing or a transient error; sign-in denies ("couldn't
 *             confirm"), the session recheck keeps the session for at most
 *             ACCESS_GRACE_MS since the last successful check.
 */
export type LauncherAccess = true | false | "error";

export async function launcherAccess(
  profileId: string,
  supabase: SupabaseClient = createAdminClient(),
): Promise<LauncherAccess> {
  try {
    const { data, error } = await supabase.rpc("launcher_can_open_app", {
      p_profile_id: profileId,
      p_audience: LAUNCHER_AUDIENCE,
    });
    if (error) {
      console.warn(`[launcher] launcher_can_open_app failed: ${error.code ?? ""} ${error.message}`);
      return "error";
    }
    return data === true;
  } catch (err) {
    console.warn(
      `[launcher] launcher_can_open_app threw: ${err instanceof Error ? err.message : String(err)}`,
    );
    return "error";
  }
}

/**
 * Profile ids (active, with launcher access to this app). Used to limit the
 * in-app "View As" picker. Returns null when the launcher can't answer —
 * callers must then treat nobody as previewable (fail closed).
 */
export async function launcherAppMembers(
  supabase: SupabaseClient = createAdminClient(),
): Promise<Set<string> | null> {
  try {
    const { data, error } = await supabase.rpc("launcher_app_members", {
      p_audience: LAUNCHER_AUDIENCE,
    });
    if (error) {
      console.warn(`[launcher] launcher_app_members failed: ${error.code ?? ""} ${error.message}`);
      return null;
    }
    const ids = new Set<string>();
    for (const row of (data ?? []) as unknown[]) {
      if (typeof row === "string") ids.add(row);
      else if (row && typeof row === "object") {
        const v = Object.values(row as Record<string, unknown>)[0];
        if (typeof v === "string") ids.add(v);
      }
    }
    return ids;
  } catch (err) {
    console.warn(
      `[launcher] launcher_app_members threw: ${err instanceof Error ? err.message : String(err)}`,
    );
    return null;
  }
}
