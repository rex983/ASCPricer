import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { generateQuotePdf } from "@/lib/pdf/generate";
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

  const pdfBuffer = await generateQuotePdf(quote as Quote);

  return new NextResponse(new Uint8Array(pdfBuffer), {
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition": `inline; filename="${quote.quote_number}.pdf"`,
    },
  });
}
