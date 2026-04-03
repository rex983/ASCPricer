import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";

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

  const supabase = createAdminClient();

  const { data, error } = await supabase
    .from("asc_pricing_data")
    .select(
      "id, version, spreadsheet_type, is_current, created_at, upload_id, asc_uploads(filename)"
    )
    .eq("region_id", regionId)
    .order("version", { ascending: false });

  if (error) {
    return NextResponse.json({ error: "Failed to load versions" }, { status: 500 });
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
