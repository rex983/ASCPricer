import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { parseSpreadsheet } from "@/lib/excel/parser";

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
function isValidUuid(v: unknown): v is string {
  return typeof v === "string" && UUID_RE.test(v);
}

export async function POST(req: NextRequest) {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const formData = await req.formData();
  const file = formData.get("file") as File | null;
  const regionId = formData.get("regionId") as string | null;

  if (!file) {
    return NextResponse.json({ error: "No file provided" }, { status: 400 });
  }
  if (!regionId) {
    return NextResponse.json(
      { error: "No region selected" },
      { status: 400 }
    );
  }

  const supabase = createAdminClient();

  // Create upload record
  const { data: upload, error: uploadError } = await supabase
    .from("asc_uploads")
    .insert({
      region_id: regionId,
      uploaded_by: isValidUuid(session.user.profileId) ? session.user.profileId : null,
      filename: file.name,
      spreadsheet_type: "standard", // will be updated after detection
      status: "processing",
    })
    .select()
    .single();

  if (uploadError) {
    return NextResponse.json(
      { error: "Failed to create upload record: " + uploadError.message },
      { status: 500 }
    );
  }

  try {
    // Parse the spreadsheet
    const buffer = await file.arrayBuffer();
    const result = parseSpreadsheet(new Uint8Array(buffer));

    if (!result.validation.valid) {
      await supabase
        .from("asc_uploads")
        .update({
          status: "failed",
          error_message: result.validation.errors.join("; "),
        })
        .eq("id", upload.id);

      return NextResponse.json(
        {
          error: "Validation failed",
          errors: result.validation.errors,
          warnings: result.validation.warnings,
        },
        { status: 422 }
      );
    }

    // Update upload with detected type and sheet count
    await supabase
      .from("asc_uploads")
      .update({
        spreadsheet_type: result.detection.type,
        sheet_count: result.detection.sheetCount,
        status: "success",
      })
      .eq("id", upload.id);

    // Deactivate previous current pricing for this region+type
    await supabase
      .from("asc_pricing_data")
      .update({ is_current: false })
      .eq("region_id", regionId)
      .eq("spreadsheet_type", result.detection.type)
      .eq("is_current", true);

    // Get next version number
    const { data: prevVersions } = await supabase
      .from("asc_pricing_data")
      .select("version")
      .eq("region_id", regionId)
      .eq("spreadsheet_type", result.detection.type)
      .order("version", { ascending: false })
      .limit(1);

    const nextVersion = (prevVersions?.[0]?.version ?? 0) + 1;

    // Insert new pricing data
    const { data: pricingData, error: pricingError } = await supabase
      .from("asc_pricing_data")
      .insert({
        region_id: regionId,
        version: nextVersion,
        is_current: true,
        spreadsheet_type: result.detection.type,
        matrices: result.matrices,
        upload_id: upload.id,
      })
      .select()
      .single();

    if (pricingError) {
      await supabase
        .from("asc_uploads")
        .update({
          status: "failed",
          error_message: "Failed to store pricing data: " + pricingError.message,
        })
        .eq("id", upload.id);

      return NextResponse.json(
        { error: "Failed to store pricing data: " + pricingError.message },
        { status: 500 }
      );
    }

    return NextResponse.json({
      success: true,
      upload: {
        id: upload.id,
        filename: file.name,
        type: result.detection.type,
        states: result.detection.states,
        sheetCount: result.detection.sheetCount,
        version: nextVersion,
      },
      pricingDataId: pricingData.id,
      validation: {
        warnings: result.validation.warnings,
      },
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unknown error";
    await supabase
      .from("asc_uploads")
      .update({ status: "failed", error_message: message })
      .eq("id", upload.id);

    return NextResponse.json(
      { error: "Parse failed: " + message },
      { status: 500 }
    );
  }
}
