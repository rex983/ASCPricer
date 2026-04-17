import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { calculatePrice } from "@/lib/pricing/engine";
import { getImpersonationContext } from "@/lib/impersonation";
import type { BuildingConfig, PricingMatrices } from "@/types/pricing";

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export async function GET(req: NextRequest) {
  const ctx = await getImpersonationContext();
  if (!ctx) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const supabase = createAdminClient();
  const { role, profileId, office } = ctx.effective;
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
  if (role === "sales_rep" || role === "bst") {
    query = query.eq("created_by", profileId);
  } else if (role === "manager" && office) {
    query = query.eq("office", office);
  }
  // admin sees all

  if (status && status !== "all") {
    query = query.eq("status", status);
  }
  if (search) {
    // Escape special PostgREST filter characters to prevent filter injection
    const s = search.replace(/[%_\\(),."']/g, "");
    if (s) {
      query = query.or(
        `quote_number.ilike.%${s}%,customer_name.ilike.%${s}%`
      );
    }
  }

  const { data, error } = await query;

  if (error) {
    console.error("Quotes list error:", error.message);
    return NextResponse.json({ error: "Failed to load quotes" }, { status: 500 });
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

  if (!UUID_RE.test(regionId)) {
    return NextResponse.json({ error: "Invalid region ID" }, { status: 400 });
  }

  // Validate office if provided
  const VALID_OFFICES = ["Harbor", "Marion"];
  if (office && !VALID_OFFICES.includes(office)) {
    return NextResponse.json({ error: "Invalid office" }, { status: 400 });
  }

  // Validate BuildingConfig numeric bounds
  const { width, length, height, taxRate, windRating } = config;
  if (typeof width !== "number" || width < 12 || width > 60) {
    return NextResponse.json({ error: "Invalid width" }, { status: 400 });
  }
  if (typeof length !== "number" || length < 20 || length > 200) {
    return NextResponse.json({ error: "Invalid length" }, { status: 400 });
  }
  if (typeof height !== "number" || height < 6 || height > 20) {
    return NextResponse.json({ error: "Invalid height" }, { status: 400 });
  }
  if (typeof taxRate !== "number" || taxRate < 0 || taxRate > 0.2) {
    return NextResponse.json({ error: "Invalid tax rate" }, { status: 400 });
  }
  if (typeof windRating !== "number" || windRating < 0 || windRating > 300) {
    return NextResponse.json({ error: "Invalid wind rating" }, { status: 400 });
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

  // Generate quote number + resolve profile UUID in parallel
  const profileId = session.user.profileId;
  const isValidUuid = !!(profileId && UUID_RE.test(profileId));

  const [quoteNumResult, profileResult] = await Promise.all([
    supabase.rpc("next_quote_number"),
    isValidUuid
      ? supabase.from("profiles").select("id").eq("id", profileId).single()
      : session.user.email
        ? supabase.from("profiles").select("id").eq("email", session.user.email).single()
        : Promise.resolve({ data: null }),
  ]);

  if (quoteNumResult.error) {
    return NextResponse.json(
      { error: "Failed to generate quote number" },
      { status: 500 }
    );
  }
  const quoteNum = quoteNumResult.data;
  const validUuid = profileResult.data?.id ?? null;

  // Resolve office: explicit param > session office > null
  const quoteOffice = office || session.user.office || null;

  const validUntil = new Date();
  validUntil.setDate(validUntil.getDate() + 30);

  // Resolve customer: use existing customerId, or auto-create from form fields
  let resolvedCustomerId: string | null = customerId || null;
  let customerName = customer?.name || null;
  let customerEmail = customer?.email || null;
  let customerPhone = customer?.phone || null;
  let customerAddress = customer?.address || null;
  let customerCity = customer?.city || null;
  let customerState = customer?.state || null;
  let customerZip = customer?.zip || null;

  if (resolvedCustomerId) {
    // Existing customer — snapshot their info onto the quote
    const { data: cust } = await supabase
      .from("asc_customers")
      .select("name, email, phone, address, city, state, zip")
      .eq("id", resolvedCustomerId)
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
  } else if (customerName) {
    // No existing customer — create one automatically
    const { data: newCust, error: custError } = await supabase
      .from("asc_customers")
      .insert({
        name: customerName,
        email: customerEmail,
        phone: customerPhone,
        address: customerAddress,
        city: customerCity,
        state: customerState,
        zip: customerZip,
        office: quoteOffice || "Harbor",
        created_by: validUuid,
      })
      .select("id")
      .single();
    if (custError) {
      console.error("Auto-create customer error:", custError.message);
    } else if (newCust) {
      resolvedCustomerId = newCust.id;
    }
  }

  const { data: quote, error: insertError } = await supabase
    .from("asc_quotes")
    .insert({
      quote_number: quoteNum,
      region_id: regionId,
      pricing_data_id: pricingRow.id,
      created_by: validUuid,
      customer_id: resolvedCustomerId,
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
    console.error("Quote insert error:", insertError.message);
    return NextResponse.json({ error: "Failed to create quote" }, { status: 500 });
  }

  return NextResponse.json(quote, { status: 201 });
}
