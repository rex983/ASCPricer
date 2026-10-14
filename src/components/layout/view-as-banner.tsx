import { Eye } from "lucide-react";
import { auth } from "@/auth";
import { LAUNCHER_URL } from "@/lib/launcher-access";
import { getLauncherRoles } from "@/lib/launcher-org";
import { ViewAsExitButton } from "./view-as-exit-button";

/**
 * Launcher admin "View as" banner. Shown on every authenticated page while
 * session.viewAs is set; all writes are refused server-side meanwhile.
 */
export async function ViewAsBanner() {
  const session = await auth();
  const viewAs = session?.viewAs;
  if (!session?.user || !viewAs) return null;

  const name = session.user.name || session.user.email || "this user";
  const roleName = String(session.user.role || "");
  const role =
    (await getLauncherRoles()).find((r) => r.name === roleName)?.displayName ||
    roleName.replace(/_/g, " ") ||
    "no role";

  return (
    <>
      <div
        role="status"
        className="fixed inset-x-0 top-0 z-[100] flex flex-wrap items-center justify-center gap-x-3 gap-y-1 bg-red-700 px-4 py-2 text-sm font-medium text-white shadow-lg"
      >
        <Eye className="h-4 w-4 shrink-0" aria-hidden />
        <span>
          Viewing as <strong>{name}</strong> ({role}) — read-only. Signed in as{" "}
          <strong>{viewAs.actorName}</strong>.
        </span>
        <ViewAsExitButton launcherUrl={LAUNCHER_URL} />
      </div>
      {/* Spacer so the fixed banner doesn't cover page content. */}
      <div aria-hidden className="h-10 shrink-0" />
    </>
  );
}
