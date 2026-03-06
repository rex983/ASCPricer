import { NextRequest, NextResponse } from "next/server";
import { createAdminClient } from "@/lib/supabase/admin";

export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ regionId: string }> }
) {
  const { regionId } = await params;
  const supabase = createAdminClient();

  const { data, error } = await supabase
    .from("asc_pricing_data")
    .select("id, region_id, spreadsheet_type, matrices, version")
    .eq("region_id", regionId)
    .eq("is_current", true)
    .single();

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
