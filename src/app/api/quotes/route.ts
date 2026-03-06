import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { calculatePrice } from "@/lib/pricing/engine";
import type { BuildingConfig, PricingMatrices } from "@/types/pricing";

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export async function GET(req: NextRequest) {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const supabase = createAdminClient();
  const url = req.nextUrl;
  const status = url.searchParams.get("status");
  const search = url.searchParams.get("search");

  let query = supabase
    .from("asc_quotes")
    .select("id, quote_number, status, customer_name, customer_state, subtotal, total, created_at, updated_at, region_id")
    .order("created_at", { ascending: false })
    .limit(100);

  if (status && status !== "all") {
    query = query.eq("status", status);
  }
  if (search) {
    query = query.or(
      `quote_number.ilike.%${search}%,customer_name.ilike.%${search}%`
    );
  }

  const { data, error } = await query;

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  return NextResponse.json(data);
}

export async function POST(req: NextRequest) {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const body = await req.json();
  const {
    regionId,
    pricingDataId,
    config,
    customer,
    notes,
  } = body as {
    regionId: string;
    pricingDataId: string | null;
    config: BuildingConfig;
    customer?: {
      name?: string;
      email?: string;
      phone?: string;
      address?: string;
      city?: string;
      state?: string;
      zip?: string;
    };
    notes?: string;
  };

  if (!regionId || !config) {
    return NextResponse.json({ error: "regionId and config are required" }, { status: 400 });
  }

  const supabase = createAdminClient();

  // Fetch matrices to server-validate pricing
  const { data: pricingRow } = await supabase
    .from("asc_pricing_data")
    .select("id, matrices")
    .eq("region_id", regionId)
    .eq("is_current", true)
    .single();

  if (!pricingRow) {
    return NextResponse.json({ error: "No pricing data for this region" }, { status: 400 });
  }

  // Server-side recalculation
  const matrices = pricingRow.matrices as PricingMatrices;
  const pricing = calculatePrice(config, matrices);

  // Generate quote number
  const { data: quoteNum, error: seqError } = await supabase.rpc("next_quote_number");
  if (seqError) {
    return NextResponse.json({ error: "Failed to generate quote number" }, { status: 500 });
  }

  const profileId = session.user.profileId;
  const validUuid = UUID_RE.test(profileId || "") ? profileId : null;

  const validUntil = new Date();
  validUntil.setDate(validUntil.getDate() + 30);

  const { data: quote, error: insertError } = await supabase
    .from("asc_quotes")
    .insert({
      quote_number: quoteNum,
      region_id: regionId,
      pricing_data_id: pricingRow.id,
      created_by: validUuid,
      status: "draft",
      customer_name: customer?.name || null,
      customer_email: customer?.email || null,
      customer_phone: customer?.phone || null,
      customer_address: customer?.address || null,
      customer_city: customer?.city || null,
      customer_state: customer?.state || null,
      customer_zip: customer?.zip || null,
      config,
      pricing,
      subtotal: pricing.subtotal,
      tax_rate: pricing.taxRate,
      tax_amount: pricing.taxAmount,
      total: pricing.total,
      notes: notes || null,
      valid_until: validUntil.toISOString(),
    })
    .select()
    .single();

  if (insertError) {
    return NextResponse.json({ error: insertError.message }, { status: 500 });
  }

  return NextResponse.json(quote, { status: 201 });
}
