import { NextResponse } from "next/server";
import type { Session } from "next-auth";
import { auth } from "@/auth";

/** Message for any write attempted during a launcher admin "View as" session. */
export function readOnlyMessage(name: string | null | undefined): string {
  return `Read-only: viewing as ${name || "another user"}`;
}

/**
 * Route-handler guard (defense in depth behind the middleware block): returns
 * a 403 response while the session is a launcher "View as" session, else null.
 *
 *   const blocked = await rejectIfViewingAs();
 *   if (blocked) return blocked;
 */
export async function rejectIfViewingAs(
  session?: Session | null,
): Promise<NextResponse | null> {
  const s = session === undefined ? await auth() : session;
  if (!s?.viewAs) return null;
  return NextResponse.json(
    { error: readOnlyMessage(s.user?.name ?? s.user?.email) },
    { status: 403 },
  );
}

/** Throwing variant for server actions / non-route code paths. */
export async function assertNotViewingAs(session?: Session | null): Promise<void> {
  const s = session === undefined ? await auth() : session;
  if (s?.viewAs) throw new Error(readOnlyMessage(s.user?.name ?? s.user?.email));
}
