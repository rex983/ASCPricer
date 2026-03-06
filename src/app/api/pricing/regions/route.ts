import { NextRequest, NextResponse } from "next/server";
import { createAdminClient } from "@/lib/supabase/admin";

export async function GET(req: NextRequest) {
  const supabase = createAdminClient();
  const type = req.nextUrl.searchParams.get("type"); // "standard" | "widespan"

  let query = supabase
    .from("asc_regions")
    .select("*")
    .eq("is_active", true)
    .order("name");

  if (type) {
    query = query.eq("spreadsheet_type", type);
  }

  const { data, error } = await query;

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  return NextResponse.json(data);
}
