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
  const { role, profileId, office } = session.user;
  const url = req.nextUrl;
  const status = url.searchParams.get("status");
  const search = url.searchParams.get("search");

  let query = supabase
    .from("asc_quotes")
    .select(
      "id, quote_number, status, customer_name, customer_state, subtotal, total, created_at, updated_at, region_id, office, created_by"
    )
    .order("created_at", { ascending: false })
    .limit(200);

  // Role-based filtering
  if (role === "sales_rep" || role === "viewer") {
    query = query.eq("created_by", profileId);
  } else if (role === "manager" && office) {
    query = query.eq("office", office);
  }
  // admin sees all

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
    customerId,
    office,
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
    customerId?: string;
    office?: string;
    notes?: string;
  };

  if (!regionId || !config) {
    return NextResponse.json(
      { error: "regionId and config are required" },
      { status: 400 }
    );
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
    return NextResponse.json(
      { error: "No pricing data for this region" },
      { status: 400 }
    );
  }

  // Server-side recalculation
  const matrices = pricingRow.matrices as PricingMatrices;
  let pricing: ReturnType<typeof calculatePrice>;
  try {
    pricing = calculatePrice(config, matrices);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    console.error("Quote pricing error:", msg);
    return NextResponse.json(
      { error: `Pricing calculation failed: ${msg}` },
      { status: 500 }
    );
  }

  // Generate quote number
  const { data: quoteNum, error: seqError } = await supabase.rpc(
    "next_quote_number"
  );
  if (seqError) {
    return NextResponse.json(
      { error: "Failed to generate quote number" },
      { status: 500 }
    );
  }

  // Validate created_by exists in profiles to avoid FK violation
  const profileId = session.user.profileId;
  let validUuid: string | null = null;
  if (profileId && UUID_RE.test(profileId)) {
    const { data: profileExists } = await supabase
      .from("profiles")
      .select("id")
      .eq("id", profileId)
      .single();
    if (profileExists) validUuid = profileId;
  }

  // Resolve office: explicit param > session office > null
  const quoteOffice = office || session.user.office || null;

  const validUntil = new Date();
  validUntil.setDate(validUntil.getDate() + 30);

  // If a customerId is provided, snapshot the customer info onto the quote
  let customerName = customer?.name || null;
  let customerEmail = customer?.email || null;
  let customerPhone = customer?.phone || null;
  let customerAddress = customer?.address || null;
  let customerCity = customer?.city || null;
  let customerState = customer?.state || null;
  let customerZip = customer?.zip || null;

  if (customerId) {
    const { data: cust } = await supabase
      .from("asc_customers")
      .select("name, email, phone, address, city, state, zip")
      .eq("id", customerId)
      .single();
    if (cust) {
      customerName = cust.name;
      customerEmail = cust.email;
      customerPhone = cust.phone;
      customerAddress = cust.address;
      customerCity = cust.city;
      customerState = cust.state;
      customerZip = cust.zip;
    }
  }

  const { data: quote, error: insertError } = await supabase
    .from("asc_quotes")
    .insert({
      quote_number: quoteNum,
      region_id: regionId,
      pricing_data_id: pricingRow.id,
      created_by: validUuid,
      customer_id: customerId || null,
      office: quoteOffice,
      status: "draft",
      customer_name: customerName,
      customer_email: customerEmail,
      customer_phone: customerPhone,
      customer_address: customerAddress,
      customer_city: customerCity,
      customer_state: customerState,
      customer_zip: customerZip,
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
