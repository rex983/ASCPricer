import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import {
  applyInsulationRates,
  getInsulationRates,
  getPricingRow,
} from "@/lib/pricing/data";

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export async function GET(
  req: NextRequest,
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

  const versionId = req.nextUrl.searchParams.get("versionId");
  if (versionId && !UUID_RE.test(versionId)) {
    return NextResponse.json({ error: "Invalid version ID" }, { status: 400 });
  }

  let data;
  try {
    data = await getPricingRow(regionId, versionId);
  } catch {
    return NextResponse.json({ error: "Failed to load pricing data" }, { status: 500 });
  }

  if (!data) {
    return NextResponse.json(
      { error: "No pricing data found for this region. Upload a spreadsheet first." },
      { status: 404 }
    );
  }

  const rates = await getInsulationRates();
  return NextResponse.json({ ...data, matrices: applyInsulationRates(data.matrices, rates) });
}
