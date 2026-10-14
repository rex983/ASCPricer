import type { SupabaseClient } from "@supabase/supabase-js";
import { createAdminClient } from "@/lib/supabase/admin";

// Must match launcher_sso_configs.jwt_audience / launcher_apps for this app.
export const LAUNCHER_AUDIENCE = "asc-pricer";

// Where users and access are managed. Used for "Managed in the BBD Launcher"
// links and the view-as Exit button.
export const LAUNCHER_URL = (
  process.env.LAUNCHER_URL || "https://bbd-launcher.vercel.app"
).replace(/\/+$/, "");

// Error code put on /login?error=... when the launcher denies access.
export const NO_LAUNCHER_ACCESS = "NoLauncherAccess";

// RPC not deployed yet (function missing / schema cache stale).
const MISSING_RPC_CODES = new Set(["PGRST202", "42883", "PGRST204"]);

let warned = false;
function warnOnce(msg: string) {
  if (warned) return;
  warned = true;
  console.warn(msg);
}

/**
 * Ask the launcher whether this profile may open this app
 * (launcher_role_app_access + launcher_user_access + launcher_apps.offices).
 *
 * Returns false only on an explicit "no". NULL (app not registered yet),
 * a missing RPC, or a transient error all fail OPEN (true) with a single
 * warning, so a deploy ahead of the launcher can't lock everyone out.
 */
export async function launcherCanOpenApp(
  profileId: string,
  supabase: SupabaseClient = createAdminClient(),
): Promise<boolean> {
  try {
    const { data, error } = await supabase.rpc("launcher_can_open_app", {
      p_profile_id: profileId,
      p_audience: LAUNCHER_AUDIENCE,
    });
    if (error) {
      if (MISSING_RPC_CODES.has(error.code ?? "")) {
        warnOnce(
          "launcher_can_open_app RPC missing — launcher access check disabled (fail open)",
        );
      } else {
        warnOnce(`launcher_can_open_app failed (fail open): ${error.message}`);
      }
      return true;
    }
    if (data === null || data === undefined) {
      warnOnce(
        `No launcher app registered with audience "${LAUNCHER_AUDIENCE}" — access check skipped (fail open)`,
      );
      return true;
    }
    return data === true;
  } catch (err) {
    warnOnce(
      `launcher_can_open_app threw (fail open): ${err instanceof Error ? err.message : String(err)}`,
    );
    return true;
  }
}
