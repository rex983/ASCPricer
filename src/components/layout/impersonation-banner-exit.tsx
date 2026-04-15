"use client";

import { useState } from "react";
import { Loader2, X } from "lucide-react";

export function ImpersonationBannerExit() {
  const [exiting, setExiting] = useState(false);

  const exit = async () => {
    setExiting(true);
    try {
      await fetch("/api/view-as", { method: "DELETE" });
      window.location.href = "/dashboard";
    } catch {
      setExiting(false);
    }
  };

  return (
    <button
      type="button"
      onClick={exit}
      disabled={exiting}
      className="flex items-center gap-1 rounded border border-amber-500/60 bg-amber-200/60 px-2 py-1 text-xs font-medium text-amber-900 hover:bg-amber-200 disabled:opacity-50 dark:border-amber-700 dark:bg-amber-900/40 dark:text-amber-100 dark:hover:bg-amber-900/60"
    >
      {exiting ? (
        <Loader2 className="h-3 w-3 animate-spin" />
      ) : (
        <X className="h-3 w-3" />
      )}
      Exit view-as
    </button>
  );
}
