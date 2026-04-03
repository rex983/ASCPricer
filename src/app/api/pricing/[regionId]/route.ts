import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";

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

  const supabase = createAdminClient();

  let query = supabase
    .from("asc_pricing_data")
    .select("id, region_id, spreadsheet_type, matrices, version, created_at")
    .eq("region_id", regionId);

  if (versionId) {
    query = query.eq("id", versionId);
  } else {
    query = query.eq("is_current", true);
  }

  const { data, error } = await query.single();

  if (error) {
    if (error.code === "PGRST116") {
      return NextResponse.json(
        { error: "No pricing data found for this region. Upload a spreadsheet first." },
        { status: 404 }
      );
    }
    return NextResponse.json({ error: "Failed to load pricing data" }, { status: 500 });
  }

  return NextResponse.json(data);
}
