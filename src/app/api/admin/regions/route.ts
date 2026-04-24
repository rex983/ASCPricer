import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { logAudit } from "@/lib/audit";

const VALID_STATES = new Set([
  "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN",
  "IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV",
  "NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN",
  "TX","UT","VT","VA","WA","WV","WI","WY",
]);

/** GET /api/admin/regions — list ALL regions (active + inactive) with pricing status */
export async function GET() {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const supabase = createAdminClient();

  // Fetch all regions
  const { data: regions, error } = await supabase
    .from("asc_regions")
    .select("*")
    .order("name");

  if (error) {
    console.error("regions GET error:", error);
    return NextResponse.json({ error: "Database operation failed" }, { status: 500 });
  }

  // Fetch current pricing data status per region+type
  const { data: pricingStatus } = await supabase
    .from("asc_pricing_data")
    .select("region_id, spreadsheet_type, version, created_at")
    .eq("is_current", true);

  // Fetch latest upload per region
  const { data: uploads } = await supabase
    .from("asc_uploads")
    .select("region_id, spreadsheet_type, filename, status, created_at")
    .eq("status", "success")
    .order("created_at", { ascending: false });

  // Build a map of region pricing info
  const pricingMap: Record<string, { version: number; uploadedAt: string }> = {};
  for (const p of pricingStatus ?? []) {
    pricingMap[p.region_id] = {
      version: p.version,
      uploadedAt: p.created_at,
    };
  }

  // Build map of last upload per region
  const uploadMap: Record<string, { filename: string; uploadedAt: string }> = {};
  for (const u of uploads ?? []) {
    if (!uploadMap[u.region_id]) {
      uploadMap[u.region_id] = { filename: u.filename, uploadedAt: u.created_at };
    }
  }

  // Combine data
  const enriched = (regions ?? []).map((r) => ({
    ...r,
    currentPricing: pricingMap[r.id] || null,
    lastUpload: uploadMap[r.id] || null,
  }));

  return NextResponse.json(enriched);
}

/** POST /api/admin/regions — create a new region */
export async function POST(req: NextRequest) {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const body = await req.json();
  const { name, states, spreadsheet_type } = body as {
    name: string;
    states: string[];
    spreadsheet_type: string;
  };

  if (!name?.trim()) {
    return NextResponse.json({ error: "Name is required" }, { status: 400 });
  }
  if (!spreadsheet_type || !["standard", "widespan"].includes(spreadsheet_type)) {
    return NextResponse.json({ error: "Spreadsheet type must be 'standard' or 'widespan'" }, { status: 400 });
  }
  if (!Array.isArray(states) || states.length === 0) {
    return NextResponse.json({ error: "At least one state is required" }, { status: 400 });
  }
  const invalidStates = states.filter((s) => !VALID_STATES.has(s));
  if (invalidStates.length > 0) {
    return NextResponse.json({ error: `Invalid state codes: ${invalidStates.join(", ")}` }, { status: 400 });
  }

  // Generate slug from states + type
  const slug = `${states.join("-").toLowerCase()}-${spreadsheet_type}`;

  const supabase = createAdminClient();

  const { data, error } = await supabase
    .from("asc_regions")
    .insert({
      name: name.trim(),
      slug,
      states,
      spreadsheet_type,
      is_active: true,
    })
    .select()
    .single();

  if (error) {
    if (error.code === "23505") {
      return NextResponse.json({ error: "A region with this combination already exists" }, { status: 409 });
    }
    console.error("regions POST error:", error);
    return NextResponse.json({ error: "Database operation failed" }, { status: 500 });
  }

  await logAudit({
    userId: session.user.profileId,
    userEmail: session.user.email,
    action: "create_region",
    resourceType: "region",
    resourceId: data.id,
    details: { name: data.name, states, spreadsheet_type },
  });

  return NextResponse.json(data, { status: 201 });
}
