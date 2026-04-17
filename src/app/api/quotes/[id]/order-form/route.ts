import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { getEffectiveUser } from "@/lib/impersonation";
import { PDFDocument } from "pdf-lib";
import fs from "fs";
import path from "path";

const WEST_COAST_STATES = ["WA", "NV", "CA", "OR"];

function fmt(n: number): string {
  return n.toFixed(2);
}

function fmtDate(iso: string): string {
  return new Date(iso).toLocaleDateString("en-US", { month: "2-digit", day: "2-digit", year: "numeric" });
}

async function fillStandardForm(pdfBytes: Buffer, quote: Record<string, unknown>, creatorName: string | null) {
  const pdf = await PDFDocument.load(pdfBytes);
  const form = pdf.getForm();

  const config = quote.config as Record<string, unknown>;
  const pricing = quote.pricing as Record<string, number>;

  const set = (name: string, value: string) => {
    try { form.getTextField(name).setText(value); } catch { /* field missing */ }
  };
  const check = (name: string) => {
    try { form.getCheckBox(name).check(); } catch { /* field missing */ }
  };
  const uncheck = (name: string) => {
    try { form.getCheckBox(name).uncheck(); } catch { /* field missing */ }
  };

  // Customer info
  set("Order Date", fmtDate(quote.created_at as string));
  set("Buyer Names", (quote.customer_name as string) || "");
  set("Buyer Address", (quote.customer_address as string) || "");
  set("City", (quote.customer_city as string) || "");
  set("State", (quote.customer_state as string) || "");
  set("Zipcode", (quote.customer_zip as string) || "");
  set("Phone Home", (quote.customer_phone as string) || "");
  set("Email_2", (quote.customer_email as string) || "");

  // Building dimensions
  set("Width", String(config.width || ""));
  set("Roof Length", String(config.length || ""));
  set("Base Length", String(config.length || ""));
  set("Gauge", String(config.gauge || ""));
  const height = Number(config.height) || 0;
  if (height > 6) set("Leg Height", String(height));

  // Base price
  set("Base Price", fmt(pricing.basePrice || 0));

  // Roof style
  uncheck("Regular Roof Style");
  uncheck("A-Frame Horizontal");
  uncheck("A-Frame Vertical");
  const roofStyle = config.roofStyle as string;
  if (roofStyle === "standard") check("Regular Roof Style");
  else if (roofStyle === "a_frame_horizontal") check("A-Frame Horizontal");
  else if (roofStyle === "a_frame_vertical") check("A-Frame Vertical");
  if (pricing.roofStyle) set("Roof Price", fmt(pricing.roofStyle));

  // Height / leg price
  if (pricing.legs) set("Height Price", fmt(pricing.legs));

  // Sides
  const sidesQty = Number(config.sidesQty) || 0;
  uncheck("Close 2 Sides");
  uncheck("Close 1 Side");
  uncheck("Horizontal Side");
  uncheck("Vertical Side");
  if (sidesQty === 2) check("Close 2 Sides");
  else if (sidesQty === 1) check("Close 1 Side");
  if (config.sidesOrientation === "vertical") check("Vertical Side");
  else if (sidesQty > 0) check("Horizontal Side");
  if (pricing.sides) set("Sides Price", fmt(pricing.sides));

  // Ends
  const endsQty = Number(config.endsQty) || 0;
  uncheck("Close 2 Ends");
  uncheck("Close 1 End");
  uncheck("Horizontal 1 End");
  uncheck("Vertical End");
  if (endsQty === 2) check("Close 2 Ends");
  else if (endsQty === 1) check("Close 1 End");
  if (config.endsOrientation === "vertical") check("Vertical End");
  else if (endsQty > 0) check("Horizontal 1 End");
  if (pricing.ends) set("Ends Price", fmt(pricing.ends));

  // Walk-in doors
  if (config.walkInDoorType && Number(config.walkInDoorQty) > 0) {
    set("walk-In door size and type", config.walkInDoorType as string);
    set("Walk-In Door QTY", String(config.walkInDoorQty));
    if (pricing.walkInDoors) set("WID Price", fmt(pricing.walkInDoors));
  }

  // Windows
  if (config.windowType && Number(config.windowQty) > 0) {
    set("Window Size 1", config.windowType as string);
    set("Window 1 QTY", String(config.windowQty));
    if (pricing.windows) set("Windows Price", fmt(pricing.windows));
  }

  // Roll-up doors (ends)
  if (config.rollUpEndSize1 && Number(config.rollUpEndQty1) > 0) {
    set("End Size 1", config.rollUpEndSize1 as string);
    set("End Qty", String(config.rollUpEndQty1));
  }
  if (config.rollUpEndSize2 && Number(config.rollUpEndQty2) > 0) {
    set("End Size 2", config.rollUpEndSize2 as string);
    set("End Qty 2", String(config.rollUpEndQty2));
  }

  // Roll-up doors (sides)
  if (config.rollUpSideSize1 && Number(config.rollUpSideQty1) > 0) {
    set("Side Size 1", config.rollUpSideSize1 as string);
    set("Side QTY 1", String(config.rollUpSideQty1));
  }
  if (config.rollUpSideSize2 && Number(config.rollUpSideQty2) > 0) {
    set("Side Size 2", config.rollUpSideSize2 as string);
    set("Side QTY 2", String(config.rollUpSideQty2));
  }

  const rudTotal = (pricing.rollUpDoorsEnds || 0) + (pricing.rollUpDoorsSides || 0);
  if (rudTotal) set("RUD Price", fmt(rudTotal));

  // Insulation
  uncheck("Roof Insulation");
  uncheck("Fully Insulated");
  if (config.insulationScope === "roof_only") check("Roof Insulation");
  else if (config.insulationScope === "fully_insulated") check("Fully Insulated");
  if (pricing.insulation) set("Insulation Price", fmt(pricing.insulation));

  // Anchors
  if (pricing.anchors) {
    check("Anchor Package");
    set("Anchor Price", fmt(pricing.anchors));
  }

  // Snow/wind engineering
  if (pricing.snowEngineering) {
    const windDetail = config.windRating ? `${config.windRating} MPH` : "";
    const snowDetail = config.snowLoad ? String(config.snowLoad) : "";
    const detail = [windDetail, snowDetail].filter(Boolean).join(" / ");
    set("OTHER2 title", "Snow/Wind Engineering");
    set("Other2", detail);
    set("Other Price 2", fmt(pricing.snowEngineering));
  }

  // Totals
  const subtotal = Number(quote.subtotal) || 0;
  const taxRate = Number(quote.tax_rate) || 0;
  const taxAmount = Number(quote.tax_amount) || 0;
  const total = Number(quote.total) || 0;

  set("Total Sale", fmt(subtotal));
  set("Tax Rate", (taxRate * 100).toFixed(2));
  set("Tax", fmt(taxAmount));
  set("Subtotal", fmt(subtotal + taxAmount));

  if (pricing.plans || pricing.calculations) {
    set("Drawings", fmt((pricing.plans || 0) + (pricing.calculations || 0)));
  }
  if (pricing.laborEquipment) set("Equipment", fmt(pricing.laborEquipment));

  set("Total", fmt(total));
  const deposit = total * 0.18;
  set("Deposit", fmt(deposit));
  set("Balance Due", fmt(total - deposit));

  // Generated by
  if (creatorName) {
    set("Other7", `Quote generated by ${creatorName}`);
  }

  return pdf.save();
}

async function fillWestCoastForm(pdfBytes: Buffer, quote: Record<string, unknown>, creatorName: string | null) {
  const pdf = await PDFDocument.load(pdfBytes);
  const form = pdf.getForm();

  const config = quote.config as Record<string, unknown>;
  const pricing = quote.pricing as Record<string, number>;

  const set = (name: string, value: string) => {
    try { form.getTextField(name).setText(value); } catch { /* field missing */ }
  };
  const check = (name: string) => {
    try { form.getCheckBox(name).check(); } catch { /* field missing */ }
  };
  const uncheck = (name: string) => {
    try { form.getCheckBox(name).uncheck(); } catch { /* field missing */ }
  };

  // Customer info
  set("Date3_af_date", fmtDate(quote.created_at as string));
  set("Buyer Names", (quote.customer_name as string) || "");
  set("Buyer Address", (quote.customer_address as string) || "");
  set("City", (quote.customer_city as string) || "");
  set("State", (quote.customer_state as string) || "");
  set("Zip", (quote.customer_zip as string) || "");
  set("Home", (quote.customer_phone as string) || "");
  set("Email", (quote.customer_email as string) || "");
  set("County", (quote.customer_city as string) || "");

  // Building dimensions
  set(" Width", String(config.width || ""));
  set(" Roof Length", String(config.length || ""));
  set(" Frame Length", String(config.length || ""));
  set("Gauge", String(config.gauge || ""));
  const height = Number(config.height) || 0;
  if (height > 6) set(" Leg Height", String(height));

  // Base price
  set("Base Price Total", fmt(pricing.basePrice || 0));

  // Roof style
  uncheck("Regular Roof Option");
  uncheck("A Frame H Roof Option");
  uncheck("A Frame V Roof Option");
  uncheck("All Vertical");
  const roofStyle = config.roofStyle as string;
  if (roofStyle === "standard") check("Regular Roof Option");
  else if (roofStyle === "a_frame_horizontal") check("A Frame H Roof Option");
  else if (roofStyle === "a_frame_vertical") check("A Frame V Roof Option");
  if (pricing.roofStyle) set("Roof Style Total", fmt(pricing.roofStyle));

  // Leg height price
  if (pricing.legs) {
    check("Leg Height");
    set("Leg Height Total", fmt(pricing.legs));
  }

  // Sides
  uncheck("Close Sides");
  uncheck("Vertical Sides");
  const sidesQty = Number(config.sidesQty) || 0;
  if (sidesQty > 0) {
    check("Close Sides");
    set("Sides Qty", String(sidesQty));
    if (config.sidesOrientation === "vertical") check("Vertical Sides");
  }
  if (pricing.sides) set("Sides Total", fmt(pricing.sides));

  // Ends
  uncheck("Closed End");
  uncheck("Vertical Ends Choice");
  uncheck("Gable End");
  const endsQty = Number(config.endsQty) || 0;
  const endType = config.endType as string;
  if (endsQty > 0) {
    if (endType === "gable" || endType === "extended_gable") check("Gable End");
    else check("Closed End");
    set("Ends Qty", String(endsQty));
    if (config.endsOrientation === "vertical") check("Vertical Ends Choice");
  }
  if (pricing.ends) set("Ends", fmt(pricing.ends));

  // Walk-in doors
  if (config.walkInDoorType && Number(config.walkInDoorQty) > 0) {
    check("Walk-In Door");
    set("Walk-In Door Size", config.walkInDoorType as string);
    set("WID Qty", String(config.walkInDoorQty));
    if (pricing.walkInDoors) set("Walk-In Door Total", fmt(pricing.walkInDoors));
  }

  // Windows
  if (config.windowType && Number(config.windowQty) > 0) {
    check("Window");
    set("Windows Qty", String(config.windowQty));
    if (pricing.windows) set("Window Total", fmt(pricing.windows));
  }

  // Roll-up doors
  if (config.rollUpEndSize1 && Number(config.rollUpEndQty1) > 0) {
    check("RUD Option");
    set("RUD Size", config.rollUpEndSize1 as string);
    set("RUD Qty", String(config.rollUpEndQty1));
  }
  const rudTotal = (pricing.rollUpDoorsEnds || 0) + (pricing.rollUpDoorsSides || 0);
  if (rudTotal) set("RUD", fmt(rudTotal));

  // Anchors
  if (pricing.anchors) {
    check("Achors Choice");
    set("Anchors", fmt(pricing.anchors));
  }

  // Snow/wind engineering
  if (pricing.snowEngineering) {
    const windDetail = config.windRating ? `${config.windRating} MPH` : "";
    const snowDetail = config.snowLoad ? String(config.snowLoad) : "";
    const detail = [windDetail, snowDetail].filter(Boolean).join(" / ");
    set("Other 1 - Specify", "Snow/Wind Eng");
    set("Other Qty 1", detail);
    set("Other 1", fmt(pricing.snowEngineering));
  }

  // Totals
  const subtotal = Number(quote.subtotal) || 0;
  const taxRate = Number(quote.tax_rate) || 0;
  const taxAmount = Number(quote.tax_amount) || 0;
  const total = Number(quote.total) || 0;

  set("TotalSale", fmt(subtotal));
  set("Tax Rate", (taxRate * 100).toFixed(2));
  set("Total Tax", fmt(taxAmount));
  set("Total Subtotal", fmt(subtotal + taxAmount));

  if (pricing.plans || pricing.calculations) {
    set("Total Drawings Price", fmt((pricing.plans || 0) + (pricing.calculations || 0)));
  }
  if (pricing.laborEquipment) set("Total Equipment Price", fmt(pricing.laborEquipment));

  set("Total Price Total", fmt(total));
  set("Total Balance Due", fmt(total));

  // Generated by
  if (creatorName) {
    set("Other 2 - Specify", `Quote by ${creatorName}`);
  }

  return pdf.save();
}

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

  // Access control
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

  // Determine which form to use based on customer state
  const customerState = ((quote.customer_state as string) || "").toUpperCase().trim();
  const isWestCoast = WEST_COAST_STATES.includes(customerState);
  const templateName = isWestCoast ? "west-coast.pdf" : "standard.pdf";

  const templatePath = path.join(process.cwd(), "public", "order-forms", templateName);
  const templateBytes = fs.readFileSync(templatePath);

  const filledPdf = isWestCoast
    ? await fillWestCoastForm(templateBytes, quote, creatorName)
    : await fillStandardForm(templateBytes, quote, creatorName);

  return new NextResponse(new Uint8Array(filledPdf), {
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition": `inline; filename="${quote.quote_number}-order-form.pdf"`,
    },
  });
}
