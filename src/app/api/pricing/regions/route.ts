import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { getActiveRegions } from "@/lib/pricing/data";

export async function GET(req: NextRequest) {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const type = req.nextUrl.searchParams.get("type"); // "standard" | "widespan"

  try {
    return NextResponse.json(await getActiveRegions(type || null));
  } catch {
    return NextResponse.json({ error: "Failed to load regions" }, { status: 500 });
  }
}
