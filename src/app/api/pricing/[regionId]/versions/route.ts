import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { getPricingVersions } from "@/lib/pricing/data";

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ regionId: string }> }
) {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { regionId } = await params;
  if (!UUID_RE.test(regionId)) {
    return NextResponse.json({ error: "Invalid region ID" }, { status: 400 });
  }

  try {
    return NextResponse.json(await getPricingVersions(regionId));
  } catch {
    return NextResponse.json({ error: "Failed to load versions" }, { status: 500 });
  }
}
