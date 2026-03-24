import { NextRequest, NextResponse } from "next/server";
import { createAdminClient } from "@/lib/supabase/admin";

export async function GET(
  req: NextRequest,
  { params }: { params: Promise<{ regionId: string }> }
) {
  const { regionId } = await params;
  const versionId = req.nextUrl.searchParams.get("versionId");
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
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  return NextResponse.json(data);
}
