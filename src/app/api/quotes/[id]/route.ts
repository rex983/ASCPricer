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

  const { id } = await params;
  const supabase = createAdminClient();

  const { data, error } = await supabase
    .from("asc_quotes")
    .select("*")
    .eq("id", id)
    .single();

  if (error) {
    if (error.code === "PGRST116") {
      return NextResponse.json({ error: "Quote not found" }, { status: 404 });
    }
    console.error("Quote operation error:", error.message);
    return NextResponse.json({ error: "Database operation failed" }, { status: 500 });
  }

  // Enforce access: sales_rep/bst can only see own quotes
  const effective = await getEffectiveUser();
  const { role, profileId, office } = effective ?? session.user;
  if (role === "sales_rep" || role === "bst") {
    if (data.created_by !== profileId) {
      return NextResponse.json({ error: "Not found" }, { status: 404 });
    }
  } else if (role === "manager" && office) {
    if (data.office && data.office !== office) {
      return NextResponse.json({ error: "Not found" }, { status: 404 });
    }
  }

  // Look up creator name
  let created_by_name: string | null = null;
  if (data.created_by) {
    const { data: profile } = await supabase
      .from("profiles")
      .select("full_name")
      .eq("id", data.created_by)
      .single();
    created_by_name = profile?.full_name ?? null;
  }

  return NextResponse.json({ ...data, created_by_name });
}

export async function DELETE(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { role, office } = session.user;
  const { id } = await params;

  // Only admins and managers may delete quotes.
  if (role !== "admin" && role !== "manager") {
    return NextResponse.json({ error: "Forbidden" }, { status: 403 });
  }

  const supabase = createAdminClient();

  // Fetch the quote to check access
  const { data: quote, error: fetchErr } = await supabase
    .from("asc_quotes")
    .select("id, created_by, office")
    .eq("id", id)
    .single();

  if (fetchErr || !quote) {
    return NextResponse.json({ error: "Quote not found" }, { status: 404 });
  }

  // Managers are still limited to their own office.
  if (role === "manager" && office) {
    if (quote.office && quote.office !== office) {
      return NextResponse.json({ error: "Forbidden" }, { status: 403 });
    }
  }

  const { error } = await supabase
    .from("asc_quotes")
    .delete()
    .eq("id", id);

  if (error) {
    console.error("Quote operation error:", error.message);
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

  // Fetch quote to check access
  const { data: existing, error: fetchErr } = await supabase
    .from("asc_quotes")
    .select("id, created_by, office")
    .eq("id", id)
    .single();

  if (fetchErr || !existing) {
    return NextResponse.json({ error: "Quote not found" }, { status: 404 });
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

  // Only allow updating specific fields
  const allowed: Record<string, unknown> = {};
  const fields = [
    "status",
    "customer_name",
    "customer_email",
    "customer_phone",
    "customer_address",
    "customer_city",
    "customer_state",
    "customer_zip",
    "customer_id",
    "notes",
    "valid_until",
  ];
  for (const f of fields) {
    if (f in body) allowed[f] = body[f];
  }

  if (Object.keys(allowed).length === 0) {
    return NextResponse.json(
      { error: "No valid fields to update" },
      { status: 400 }
    );
  }

  const { data, error } = await supabase
    .from("asc_quotes")
    .update(allowed)
    .eq("id", id)
    .select()
    .single();

  if (error) {
    console.error("Quote operation error:", error.message);
    return NextResponse.json({ error: "Database operation failed" }, { status: 500 });
  }

  return NextResponse.json(data);
}
