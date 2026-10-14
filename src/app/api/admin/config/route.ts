import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { logAudit } from "@/lib/audit";
import {
  APP_CONFIG_TAG,
  INSULATION_RATES_KEY,
  getAppConfig,
  invalidatePricingCache,
  parseInsulationRates,
} from "@/lib/pricing/data";
import { rejectIfViewingAs } from "@/lib/view-as";

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export async function GET(req: NextRequest) {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { searchParams } = new URL(req.url);
  // Only well-formed scopes reach the cache key (anything else could never
  // match a scoped row anyway, so the result is the global config).
  const rawRegionId = searchParams.get("regionId");
  const rawType = searchParams.get("spreadsheetType");
  const regionId = rawRegionId && UUID_RE.test(rawRegionId) ? rawRegionId : null;
  const spreadsheetType =
    rawType === "standard" || rawType === "widespan" ? rawType : null;

  try {
    return NextResponse.json(await getAppConfig(regionId, spreadsheetType));
  } catch {
    return NextResponse.json(
      { error: "Failed to load config" },
      { status: 500 }
    );
  }
}

export async function PUT(req: NextRequest) {
  const blocked = await rejectIfViewingAs();
  if (blocked) return blocked;

  const session = await auth();
  if (!session?.user || session.user.permRole !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const body = await req.json();
  const { key, value } = body as { key: string; value: unknown };

  if (!key || value === undefined) {
    return NextResponse.json(
      { error: "Missing key or value" },
      { status: 400 }
    );
  }

  if (key === INSULATION_RATES_KEY && !parseInsulationRates(value)) {
    return NextResponse.json(
      { error: "insulation_rates needs positive numeric fiberglass and thermal rates" },
      { status: 400 }
    );
  }

  const supabase = createAdminClient();

  // Update the global row (region_id IS NULL)
  const { error } = await supabase
    .from("asc_app_config")
    .update({
      value,
      updated_at: new Date().toISOString(),
      updated_by: session.user.email ?? session.user.profileId ?? null,
    })
    .eq("key", key)
    .is("region_id", null);

  if (error) {
    return NextResponse.json(
      { error: "Failed to update config" },
      { status: 500 }
    );
  }

  invalidatePricingCache(APP_CONFIG_TAG);

  await logAudit({
    userId: session.user.profileId,
    userEmail: session.user.email,
    action: "update_config",
    resourceType: "app_config",
    resourceId: key,
    details: { key, value },
  });

  return NextResponse.json({ success: true });
}
