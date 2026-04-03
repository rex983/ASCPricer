import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";

export async function GET(req: NextRequest) {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

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
    return NextResponse.json({ error: "Failed to load regions" }, { status: 500 });
  }

  return NextResponse.json(data);
}
