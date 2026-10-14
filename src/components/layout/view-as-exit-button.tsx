"use client";

import { useState } from "react";
import { signOut } from "next-auth/react";
import { Loader2, LogOut } from "lucide-react";

/** Sign out of this app and go back to the launcher dashboard. */
export function ViewAsExitButton({ launcherUrl }: { launcherUrl: string }) {
  const [exiting, setExiting] = useState(false);

  const exit = async () => {
    setExiting(true);
    try {
      // NextAuth only redirects same-origin, so navigate to the launcher ourselves.
      await signOut({ redirect: false });
    } finally {
      window.location.href = `${launcherUrl}/dashboard`;
    }
  };

  return (
    <button
      type="button"
      onClick={exit}
      disabled={exiting}
      className="inline-flex items-center gap-1 rounded border border-white/70 bg-white px-2 py-0.5 text-xs font-semibold text-red-700 hover:bg-red-50 disabled:opacity-60"
    >
      {exiting ? (
        <Loader2 className="h-3 w-3 animate-spin" />
      ) : (
        <LogOut className="h-3 w-3" />
      )}
      Exit
    </button>
  );
}
