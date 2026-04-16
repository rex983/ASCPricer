import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { getEffectiveUser } from "@/lib/impersonation";

export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const effective = await getEffectiveUser();
  const { role, profileId, office } = effective ?? session.user;
  const { id } = await params;
  const supabase = createAdminClient();

  // Fetch customer, quotes, and assigned rep in parallel
  const [customerResult, quotesResult, repsResult] = await Promise.all([
    supabase.from("asc_customers").select("*").eq("id", id).single(),
    supabase
      .from("asc_quotes")
      .select("id, quote_number, status, subtotal, total, created_at")
      .eq("customer_id", id)
      .order("created_at", { ascending: false }),
    supabase
      .from("asc_sales_reps")
      .select("id, name")
      .eq("is_active", true)
      .order("name"),
  ]);

  if (customerResult.error) {
    if (customerResult.error.code === "PGRST116") {
      return NextResponse.json({ error: "Customer not found" }, { status: 404 });
    }
    console.error("Customer fetch error:", customerResult.error.message);
    return NextResponse.json({ error: "Database operation failed" }, { status: 500 });
  }

  // Access control: sales_rep/bst can only see own customers, manager own office
  const customer = customerResult.data;
  if (role === "sales_rep" || role === "bst") {
    if (customer.created_by !== profileId) {
      return NextResponse.json({ error: "Not found" }, { status: 404 });
    }
  } else if (role === "manager" && office) {
    if (customer.office && customer.office !== office) {
      return NextResponse.json({ error: "Not found" }, { status: 404 });
    }
  }

  // Find assigned rep name
  const allReps = repsResult.data || [];
  const assignedRep = customer.assigned_rep_id
    ? allReps.find((r) => r.id === customer.assigned_rep_id) || null
    : null;

  return NextResponse.json({
    ...customer,
    quotes: quotesResult.data || [],
    assigned_rep: assignedRep,
    available_reps: allReps,
  });
}

export async function DELETE(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const effective = await getEffectiveUser();
  const { role, profileId, office } = effective ?? session.user;
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

  // Access control: admin=all, manager=own office, sales_rep/bst=own only
  if (role === "sales_rep" || role === "bst") {
    if (customer.created_by !== profileId) {
      return NextResponse.json({ error: "Forbidden" }, { status: 403 });
    }
  } else if (role === "manager" && office) {
    if (customer.office && customer.office !== office) {
      return NextResponse.json({ error: "Forbidden" }, { status: 403 });
    }
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
    console.error("Customer operation error:", error.message);
    return NextResponse.json({ error: "Database operation failed" }, { status: 500 });
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

  const effective = await getEffectiveUser();
  const { role, profileId, office } = effective ?? session.user;
  const { id } = await params;
  const supabase = createAdminClient();

  // Fetch customer to check access
  const { data: existing, error: fetchErr } = await supabase
    .from("asc_customers")
    .select("id, created_by, office")
    .eq("id", id)
    .single();

  if (fetchErr || !existing) {
    return NextResponse.json({ error: "Customer not found" }, { status: 404 });
  }

  // Access control: admin=all, manager=own office, sales_rep/bst=own only
  if (role === "sales_rep" || role === "bst") {
    if (existing.created_by !== profileId) {
      return NextResponse.json({ error: "Forbidden" }, { status: 403 });
    }
  } else if (role === "manager" && office) {
    if (existing.office && existing.office !== office) {
      return NextResponse.json({ error: "Forbidden" }, { status: 403 });
    }
  }

  const body = await req.json();

  const allowed: Record<string, unknown> = {};
  const fields = [
    "name", "email", "phone", "address", "city", "state", "zip", "notes",
  ];
  for (const f of fields) {
    if (f in body) allowed[f] = body[f];
  }
  if ("assigned_rep_id" in body) {
    const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    allowed.assigned_rep_id = body.assigned_rep_id && UUID_RE.test(body.assigned_rep_id)
      ? body.assigned_rep_id
      : null;
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
    console.error("Customer operation error:", error.message);
    return NextResponse.json({ error: "Database operation failed" }, { status: 500 });
  }

  return NextResponse.json(data);
}
