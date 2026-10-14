import { NextResponse } from "next/server";
import { auth } from "@/auth";
import {
  getLauncherDepartments,
  getLauncherOffices,
  getLauncherRoles,
} from "@/lib/launcher-org";

/** GET /api/org — launcher-managed offices / departments / roles for dropdowns and labels. */
export async function GET() {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }
  const [offices, departments, roles] = await Promise.all([
    getLauncherOffices(),
    getLauncherDepartments(),
    getLauncherRoles(),
  ]);
  return NextResponse.json({
    offices,
    departments,
    roles: roles.map((r) => ({ name: r.name, displayName: r.displayName })),
  });
}
