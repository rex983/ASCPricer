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
    .select(
      "id, version, spreadsheet_type, is_current, created_at, upload_id, asc_uploads(filename)"
    )
    .eq("region_id", regionId)
    .order("version", { ascending: false });

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  // Flatten the upload filename into each row
  const versions = (data || []).map((row) => {
    const upload = row.asc_uploads as unknown as { filename: string } | null;
    return {
      id: row.id,
      version: row.version,
      spreadsheetType: row.spreadsheet_type,
      isCurrent: row.is_current,
      createdAt: row.created_at,
      filename: upload?.filename || null,
    };
  });

  return NextResponse.json(versions);
}
