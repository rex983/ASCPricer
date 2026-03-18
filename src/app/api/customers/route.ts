import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";

export async function GET(req: NextRequest) {
  const session = await auth();
  if (!session?.user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const supabase = createAdminClient();
  const { role, profileId, office } = session.user;
  const search = req.nextUrl.searchParams.get("search");

  let query = supabase
    .from("asc_customers")
    .select("id, name, email, phone, city, state, office, created_at")
    .order("name", { ascending: true })
    .limit(200);

  // Role-based filtering
  if (role === "sales_rep" || role === "viewer") {
    query = query.eq("created_by", profileId);
  } else if (role === "manager" && office) {
    query = query.eq("office", office);
  }
  // admin sees all

  if (search) {
    query = query.or(
      `name.ilike.%${search}%,email.ilike.%${search}%,phone.ilike.%${search}%`
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
  const { name, email, phone, address, city, state, zip, notes, office } =
    body as {
      name: string;
      email?: string;
      phone?: string;
      address?: string;
      city?: string;
      state?: string;
      zip?: string;
      notes?: string;
      office: string;
    };

  if (!name || !office) {
    return NextResponse.json(
      { error: "name and office are required" },
      { status: 400 }
    );
  }

  if (office !== "Harbor" && office !== "Marion") {
    return NextResponse.json(
      { error: "office must be Harbor or Marion" },
      { status: 400 }
    );
  }

  const supabase = createAdminClient();
  const UUID_RE =
    /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  const createdBy = UUID_RE.test(session.user.profileId || "")
    ? session.user.profileId
    : null;

  const { data, error } = await supabase
    .from("asc_customers")
    .insert({
      name,
      email: email || null,
      phone: phone || null,
      address: address || null,
      city: city || null,
      state: state || null,
      zip: zip || null,
      notes: notes || null,
      office,
      created_by: createdBy,
    })
    .select()
    .single();

  if (error) {
    return NextResponse.json({ error: error.message }, { status: 500 });
  }

  return NextResponse.json(data, { status: 201 });
}
