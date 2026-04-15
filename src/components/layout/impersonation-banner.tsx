import { Eye } from "lucide-react";
import { getImpersonationContext } from "@/lib/impersonation";
import { ImpersonationBannerExit } from "./impersonation-banner-exit";

export async function ImpersonationBanner() {
  const ctx = await getImpersonationContext();
  if (!ctx?.isImpersonating || !ctx.target) return null;

  const name =
    ctx.target.name?.trim() || ctx.target.email.split("@")[0];

  return (
    <div className="flex items-center justify-between gap-3 border-b border-amber-400/70 bg-amber-100 px-4 py-2 text-sm text-amber-900 dark:border-amber-700 dark:bg-amber-950/40 dark:text-amber-100">
      <div className="flex items-center gap-2">
        <Eye className="h-4 w-4 shrink-0" />
        <span>
          Viewing as <strong>{name}</strong>
          {ctx.target.office && (
            <span className="ml-1 text-amber-700 dark:text-amber-300">
              ({ctx.target.office})
            </span>
          )}
          <span className="ml-2 text-xs text-amber-700 dark:text-amber-300">
            — reads are filtered to this rep. New records are still created as you.
          </span>
        </span>
      </div>
      <ImpersonationBannerExit />
    </div>
  );
}
