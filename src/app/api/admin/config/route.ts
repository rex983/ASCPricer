import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { logAudit } from "@/lib/audit";

export async function GET() {
  const supabase = createAdminClient();
  const { data, error } = await supabase
    .from("asc_app_config")
    .select("key, value");

  if (error) {
    return NextResponse.json(
      { error: "Failed to load config: " + error.message },
      { status: 500 }
    );
  }

  const config: Record<string, unknown> = {};
  for (const row of data ?? []) {
    config[row.key] = row.value;
  }

  return NextResponse.json(config);
}

export async function PUT(req: NextRequest) {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
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

  const supabase = createAdminClient();
  const { error } = await supabase
    .from("asc_app_config")
    .update({
      value,
      updated_at: new Date().toISOString(),
      updated_by: session.user.email ?? session.user.profileId ?? null,
    })
    .eq("key", key);

  if (error) {
    return NextResponse.json(
      { error: "Failed to update config: " + error.message },
      { status: 500 }
    );
  }

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
