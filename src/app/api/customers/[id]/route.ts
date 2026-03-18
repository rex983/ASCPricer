import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";

export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { id } = await params;
  const supabase = createAdminClient();

  const { data: customer, error } = await supabase
    .from("asc_customers")
    .select("*")
    .eq("id", id)
    .single();

  if (error) {
    if (error.code === "PGRST116") {
      return NextResponse.json({ error: "Customer not found" }, { status: 404 });
    }
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  // Also fetch quotes linked to this customer
  const { data: quotes } = await supabase
    .from("asc_quotes")
    .select("id, quote_number, status, subtotal, total, created_at")
    .eq("customer_id", id)
    .order("created_at", { ascending: false });

  return NextResponse.json({ ...customer, quotes: quotes || [] });
}

export async function DELETE(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { role, profileId, office } = session.user;
  const { id } = await params;
  const supabase = createAdminClient();

  // Fetch the customer to check access
  const { data: customer, error: fetchErr } = await supabase
    .from("asc_customers")
    .select("id, created_by, office")
    .eq("id", id)
    .single();

  if (fetchErr || !customer) {
    return NextResponse.json({ error: "Customer not found" }, { status: 404 });
  }

  // Access control: admin=all, manager=own office, sales_rep=own only
  if (role === "sales_rep") {
    if (customer.created_by !== profileId) {
      return NextResponse.json({ error: "Forbidden" }, { status: 403 });
    }
  } else if (role === "manager" && office) {
    if (customer.office && customer.office !== office) {
      return NextResponse.json({ error: "Forbidden" }, { status: 403 });
    }
  } else if (role === "viewer") {
    return NextResponse.json({ error: "Forbidden" }, { status: 403 });
  }

  // Nullify customer_id on any linked quotes (don't cascade-delete quotes)
  await supabase
    .from("asc_quotes")
    .update({ customer_id: null })
    .eq("customer_id", id);

  const { error } = await supabase
    .from("asc_customers")
    .delete()
    .eq("id", id);

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  return NextResponse.json({ success: true });
}

export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { id } = await params;
  const body = await req.json();
  const supabase = createAdminClient();

  const allowed: Record<string, unknown> = {};
  const fields = [
    "name", "email", "phone", "address", "city", "state", "zip", "notes",
  ];
  for (const f of fields) {
    if (f in body) allowed[f] = body[f];
  }

  if (Object.keys(allowed).length === 0) {
    return NextResponse.json({ error: "No valid fields" }, { status: 400 });
  }

  const { data, error } = await supabase
    .from("asc_customers")
    .update(allowed)
    .eq("id", id)
    .select()
    .single();

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  return NextResponse.json(data);
}
