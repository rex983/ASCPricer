import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { generateQuotePdf } from "@/lib/pdf/generate";
import { getEffectiveUser } from "@/lib/impersonation";
import type { Quote } from "@/types/quote";

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

  const { data: quote, error } = await supabase
    .from("asc_quotes")
    .select("*")
    .eq("id", id)
    .single();

  if (error || !quote) {
    return NextResponse.json({ error: "Quote not found" }, { status: 404 });
  }

  // Access control: sales_rep/bst can only see own quotes, manager own office
  const effective = await getEffectiveUser();
  const { role, profileId, office } = effective ?? session.user;
  if (role === "sales_rep" || role === "bst") {
    if (quote.created_by !== profileId) {
      return NextResponse.json({ error: "Not found" }, { status: 404 });
    }
  } else if (role === "manager" && office) {
    if (quote.office && quote.office !== office) {
      return NextResponse.json({ error: "Not found" }, { status: 404 });
    }
  }

  // Look up creator name
  let creatorName: string | null = null;
  if (quote.created_by) {
    const { data: profile } = await supabase
      .from("profiles")
      .select("full_name")
      .eq("id", quote.created_by)
      .single();
    creatorName = profile?.full_name ?? null;
  }

  const pdfBuffer = await generateQuotePdf(quote as Quote, creatorName);

  return new NextResponse(new Uint8Array(pdfBuffer), {
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition": `inline; filename="${quote.quote_number}.pdf"`,
    },
  });
}
