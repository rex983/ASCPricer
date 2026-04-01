import React from "react";
import fs from "fs";
import path from "path";
import {
  Document,
  Page,
  Text,
  View,
  Image,
  StyleSheet,
  renderToBuffer,
} from "@react-pdf/renderer";
import type { Quote } from "@/types/quote";
import type { PriceBreakdown, SnowEngineeringBreakdown } from "@/types/pricing";

// Load logo as base64 data URI for embedding in PDF
function getLogoDataUri(): string {
  try {
    const logoPath = path.join(process.cwd(), "public", "ASCI.png");
    const buf = fs.readFileSync(logoPath);
    return `data:image/png;base64,${buf.toString("base64")}`;
  } catch {
    return "";
  }
}

const logoDataUri = getLogoDataUri();

const styles = StyleSheet.create({
  page: { padding: 40, fontSize: 10, fontFamily: "Helvetica" },
  header: { flexDirection: "row", justifyContent: "space-between", marginBottom: 20, alignItems: "center" },
  logo: { width: 72, height: 72 },
  title: { fontSize: 20, fontWeight: "bold", color: "#1a1a1a" },
  subtitle: { fontSize: 10, color: "#666", marginTop: 4 },
  section: { marginBottom: 16 },
  sectionTitle: { fontSize: 12, fontWeight: "bold", marginBottom: 6, color: "#333", borderBottomWidth: 1, borderBottomColor: "#ddd", paddingBottom: 4 },
  row: { flexDirection: "row", justifyContent: "space-between", paddingVertical: 2 },
  label: { color: "#555" },
  value: { fontWeight: "bold" },
  divider: { borderBottomWidth: 1, borderBottomColor: "#ccc", marginVertical: 6 },
  totalRow: { flexDirection: "row", justifyContent: "space-between", paddingVertical: 4 },
  totalLabel: { fontSize: 14, fontWeight: "bold" },
  totalValue: { fontSize: 14, fontWeight: "bold" },
  grid: { flexDirection: "row", flexWrap: "wrap" },
  gridItem: { width: "50%", paddingVertical: 2 },
  footer: { position: "absolute", bottom: 30, left: 40, right: 40, textAlign: "center", fontSize: 8, color: "#999" },
  // Snow breakdown table styles
  table: { marginTop: 8 },
  tableHeader: { flexDirection: "row", backgroundColor: "#333", paddingVertical: 6, paddingHorizontal: 8 },
  tableHeaderCell: { color: "#fff", fontSize: 9, fontWeight: "bold" },
  tableRow: { flexDirection: "row", paddingVertical: 5, paddingHorizontal: 8, borderBottomWidth: 0.5, borderBottomColor: "#ddd" },
  tableRowAlt: { flexDirection: "row", paddingVertical: 5, paddingHorizontal: 8, borderBottomWidth: 0.5, borderBottomColor: "#ddd", backgroundColor: "#f9f9f9" },
  tableCell: { fontSize: 9, color: "#333" },
  tableCellBold: { fontSize: 9, fontWeight: "bold", color: "#1a1a1a" },
  // Column widths for snow table
  colComponent: { width: "25%" },
  colOriginal: { width: "18%", textAlign: "center" as const },
  colSpacing: { width: "20%", textAlign: "center" as const },
  colExtra: { width: "17%", textAlign: "center" as const },
  colCost: { width: "20%", textAlign: "right" as const },
});

function fmt(n: number): string {
  return "$" + n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function PriceLine({ label, detail, value }: { label: string; detail?: string; value: number }) {
  if (value === 0) return null;
  return (
    <View style={styles.row}>
      <Text style={styles.label}>
        {label}
        {detail ? <Text style={{ color: "#888" }}>  {detail}</Text> : null}
      </Text>
      <Text>{fmt(value)}</Text>
    </View>
  );
}

function PageHeader({ quote }: { quote: Quote }) {
  return (
    <View style={styles.header}>
      <View style={{ flexDirection: "row", alignItems: "center", gap: 12 }}>
        {logoDataUri ? <Image src={logoDataUri} style={styles.logo} /> : null}
        <View>
          <Text style={styles.title}>American Steel Carports</Text>
          <Text style={styles.subtitle}>Building Quote</Text>
        </View>
      </View>
      <View style={{ alignItems: "flex-end" }}>
        <Text style={{ fontSize: 14, fontWeight: "bold" }}>{quote.quote_number}</Text>
        <Text style={styles.subtitle}>
          {new Date(quote.created_at).toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" })}
        </Text>
        {quote.valid_until && (
          <Text style={styles.subtitle}>
            Valid until {new Date(quote.valid_until).toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" })}
          </Text>
        )}
      </View>
    </View>
  );
}

function SnowBreakdownPage({ quote }: { quote: Quote }) {
  const p = quote.pricing as PriceBreakdown;
  const c = quote.config;
  const breakdown = p.snowEngineeringBreakdown as SnowEngineeringBreakdown | undefined;
  const components = breakdown?.components ?? [];
  const totalCost = breakdown?.totalCost ?? 0;

  return (
    <Page size="LETTER" style={styles.page}>
      {/* Header */}
      <PageHeader quote={quote} />

      {/* Engineering Summary */}
      <View style={styles.section}>
        <Text style={styles.sectionTitle}>Snow / Wind Load Engineering Breakdown</Text>
        <View style={styles.grid}>
          <View style={styles.gridItem}>
            <Text><Text style={styles.label}>Building: </Text>{c.width}&apos; x {c.length}&apos; x {c.height}&apos;</Text>
          </View>
          <View style={styles.gridItem}>
            <Text><Text style={styles.label}>Wind Rating: </Text>{c.windRating} MPH</Text>
          </View>
          <View style={styles.gridItem}>
            <Text><Text style={styles.label}>Snow Load: </Text>{c.snowLoad || "None"}</Text>
          </View>
          <View style={styles.gridItem}>
            <Text><Text style={styles.label}>State: </Text>{c.state || "N/A"}</Text>
          </View>
          <View style={styles.gridItem}>
            <Text><Text style={styles.label}>Roof Style: </Text>{c.roofStyle.replace(/_/g, " ")}</Text>
          </View>
          <View style={styles.gridItem}>
            <Text><Text style={styles.label}>Sides: </Text>{c.sidesCoverage.replace(/_/g, " ")} x{c.sidesQty}</Text>
          </View>
        </View>
      </View>

      {/* Component Breakdown Table */}
      <View style={styles.section}>
        <Text style={{ fontSize: 11, fontWeight: "bold", marginBottom: 8, color: "#333" }}>
          Per-Component Breakdown
        </Text>

        {p.contactEngineer && (
          <View style={{ padding: 10, backgroundColor: "#fef3c7", borderRadius: 4, marginBottom: 8 }}>
            <Text style={{ fontSize: 10, fontWeight: "bold", color: "#92400e" }}>
              Snow/wind loads exceed standard engineering — contact engineer for custom pricing.
            </Text>
          </View>
        )}

        {!c.snowLoad && components.length === 0 && (
          <View style={{ padding: 10, backgroundColor: "#f5f5f5", borderRadius: 4, marginBottom: 8 }}>
            <Text style={{ fontSize: 10, color: "#555" }}>
              No snow load selected. Only wind rating ({c.windRating} MPH) applies.
            </Text>
          </View>
        )}

        {/* Table Header */}
        <View style={styles.tableHeader}>
          <Text style={[styles.tableHeaderCell, styles.colComponent]}>Component</Text>
          <Text style={[styles.tableHeaderCell, styles.colOriginal]}>Original Count</Text>
          <Text style={[styles.tableHeaderCell, styles.colSpacing]}>Req. Spacing</Text>
          <Text style={[styles.tableHeaderCell, styles.colExtra]}>Extra Needed</Text>
          <Text style={[styles.tableHeaderCell, styles.colCost]}>Cost</Text>
        </View>

        {/* Table Rows */}
        {components.map((comp, i) => (
          <View key={i} style={i % 2 === 1 ? styles.tableRowAlt : styles.tableRow}>
            <Text style={[styles.tableCellBold, styles.colComponent]}>{comp.name}</Text>
            <Text style={[styles.tableCell, styles.colOriginal]}>{comp.originalCount}</Text>
            <Text style={[styles.tableCell, styles.colSpacing]}>
              {comp.requiredSpacing > 0 ? `${comp.requiredSpacing}"` : "—"}
            </Text>
            <Text style={[styles.tableCell, styles.colExtra]}>
              {comp.extraNeeded > 0 ? `+${comp.extraNeeded}` : "0"}
            </Text>
            <Text style={[styles.tableCellBold, styles.colCost]}>{fmt(comp.cost)}</Text>
          </View>
        ))}

        {/* Total Row */}
        <View style={{ flexDirection: "row", paddingVertical: 8, paddingHorizontal: 8, backgroundColor: "#333", marginTop: 1 }}>
          <Text style={[styles.tableHeaderCell, styles.colComponent]}>TOTAL</Text>
          <Text style={[styles.tableHeaderCell, styles.colOriginal]}></Text>
          <Text style={[styles.tableHeaderCell, styles.colSpacing]}></Text>
          <Text style={[styles.tableHeaderCell, styles.colExtra]}></Text>
          <Text style={[styles.tableHeaderCell, styles.colCost]}>{fmt(totalCost)}</Text>
        </View>
      </View>

      {/* Verification Note */}
      <View style={{ marginTop: 12, padding: 12, backgroundColor: "#f5f5f5", borderRadius: 4 }}>
        <Text style={{ fontSize: 9, fontWeight: "bold", color: "#333", marginBottom: 4 }}>
          Engineering Notes
        </Text>
        <Text style={{ fontSize: 8, color: "#555", lineHeight: 1.5 }}>
          {`\u2022 Component costs above sum to the Snow/Wind Engineering total of ${fmt(totalCost)} on page 1.`}
        </Text>
        <Text style={{ fontSize: 8, color: "#555", lineHeight: 1.5 }}>
          {"\u2022 \"Extra Needed\" = additional components beyond standard configuration to meet load requirements."}
        </Text>
        <Text style={{ fontSize: 8, color: "#555", lineHeight: 1.5 }}>
          {"\u2022 Required spacing is the maximum allowable center-to-center distance for the given load."}
        </Text>
      </View>

      {/* Footer */}
      <Text style={styles.footer}>
        Snow / Wind Load Engineering Breakdown — {quote.quote_number}
      </Text>
    </Page>
  );
}

function QuoteDocument({ quote }: { quote: Quote }) {
  const p = quote.pricing as PriceBreakdown;
  const c = quote.config;

  return (
    <Document>
      <Page size="LETTER" style={styles.page}>
        {/* Header */}
        <PageHeader quote={quote} />

        {/* Customer */}
        {quote.customer_name && (
          <View style={styles.section}>
            <Text style={styles.sectionTitle}>Customer</Text>
            <Text>{quote.customer_name}</Text>
            {quote.customer_email && <Text style={styles.label}>{quote.customer_email}</Text>}
            {quote.customer_phone && <Text style={styles.label}>{quote.customer_phone}</Text>}
            {quote.customer_address && (
              <Text style={styles.label}>
                {quote.customer_address}
                {quote.customer_city ? `, ${quote.customer_city}` : ""}
                {quote.customer_state ? `, ${quote.customer_state}` : ""}
                {quote.customer_zip ? ` ${quote.customer_zip}` : ""}
              </Text>
            )}
          </View>
        )}

        {/* Building Config */}
        <View style={styles.section}>
          <Text style={styles.sectionTitle}>Building Configuration</Text>
          <View style={styles.grid}>
            <View style={styles.gridItem}><Text><Text style={styles.label}>Width: </Text>{c.width}&apos;</Text></View>
            <View style={styles.gridItem}><Text><Text style={styles.label}>Length: </Text>{c.length}&apos;</Text></View>
            <View style={styles.gridItem}><Text><Text style={styles.label}>Height: </Text>{c.height}&apos;</Text></View>
            <View style={styles.gridItem}><Text><Text style={styles.label}>Gauge: </Text>{c.gauge}G</Text></View>
            <View style={styles.gridItem}><Text><Text style={styles.label}>Roof: </Text>{c.roofStyle.replace(/_/g, " ")}</Text></View>
            <View style={styles.gridItem}><Text><Text style={styles.label}>Sides: </Text>{c.sidesCoverage} x{c.sidesQty}</Text></View>
            <View style={styles.gridItem}><Text><Text style={styles.label}>Ends: </Text>{c.endType} x{c.endsQty}</Text></View>
            <View style={styles.gridItem}><Text><Text style={styles.label}>Insulation: </Text>{c.insulationType}</Text></View>
          </View>
        </View>

        {/* Price Breakdown */}
        <View style={styles.section}>
          <Text style={styles.sectionTitle}>Price Breakdown</Text>
          <PriceLine label="Base Price" detail={`${c.width}x${c.length}`} value={p.basePrice} />
          <PriceLine label="Roof Style" detail={c.roofStyle.replace(/_/g, " ")} value={p.roofStyle} />
          <PriceLine label="Leg Height" detail={`${c.height}'`} value={p.legs} />
          <PriceLine label="Sides" detail={`${c.sidesCoverage.replace(/_/g, " ")} x${c.sidesQty}`} value={p.sides} />
          <PriceLine label="Ends" detail={`${c.endType} x${c.endsQty}`} value={p.ends} />
          <PriceLine label="Walk-In Doors" detail={c.walkInDoorType ? `${c.walkInDoorType} x${c.walkInDoorQty}` : undefined} value={p.walkInDoors} />
          <PriceLine label="Windows" detail={c.windowType ? `${c.windowType} x${c.windowQty}` : undefined} value={p.windows} />
          <PriceLine
            label="Roll-Up Doors (Ends)"
            detail={[
              c.rollUpEndSize1 && c.rollUpEndQty1 > 0 ? `${c.rollUpEndSize1} x${c.rollUpEndQty1}` : "",
              c.rollUpEndSize2 && c.rollUpEndQty2 > 0 ? `${c.rollUpEndSize2} x${c.rollUpEndQty2}` : "",
            ].filter(Boolean).join(", ") || undefined}
            value={p.rollUpDoorsEnds}
          />
          <PriceLine
            label="Roll-Up Doors (Sides)"
            detail={[
              c.rollUpSideSize1 && c.rollUpSideQty1 > 0 ? `${c.rollUpSideSize1} x${c.rollUpSideQty1}` : "",
              c.rollUpSideSize2 && c.rollUpSideQty2 > 0 ? `${c.rollUpSideSize2} x${c.rollUpSideQty2}` : "",
            ].filter(Boolean).join(", ") || undefined}
            value={p.rollUpDoorsSides}
          />
          <PriceLine label="Insulation" detail={c.insulationType !== "none" ? `${c.insulationType} - ${c.insulationScope?.replace(/_/g, " ")}` : undefined} value={p.insulation} />
          <PriceLine label="Wainscot" detail={c.wainscot && c.wainscot !== "none" ? c.wainscot : undefined} value={p.wainscot} />
          {p.contactEngineer ? (
            <View style={styles.row}>
              <Text style={styles.label}>Snow/Wind Engineering</Text>
              <Text style={{ color: "#d97706", fontWeight: "bold" }}>Contact Engineer</Text>
            </View>
          ) : (
            <PriceLine label="Snow/Wind Engineering" detail={[c.snowLoad ? c.snowLoad.replace(/^(\d+)\s*/, "$1 psf ") : "", c.windRating ? `${c.windRating} MPH` : ""].filter(Boolean).join(", ") || undefined} value={p.snowEngineering} />
          )}
          <PriceLine label="Diagonal Bracing" value={p.diagonalBracing} />
          <PriceLine label="Anchors" detail={c.anchorType && c.anchorType !== "none" ? c.anchorType.replace(/_/g, " ") : undefined} value={p.anchors} />

          <View style={styles.divider} />

          <View style={styles.row}>
            <Text style={{ fontWeight: "bold" }}>Subtotal</Text>
            <Text style={{ fontWeight: "bold" }}>{fmt(p.subtotal)}</Text>
          </View>
          <View style={styles.row}>
            <Text style={styles.label}>Tax ({(p.taxRate * 100).toFixed(2)}%)</Text>
            <Text>{fmt(p.taxAmount)}</Text>
          </View>
          <PriceLine label="Labor / Equipment" value={p.laborEquipment} />

          <View style={styles.divider} />

          <View style={styles.totalRow}>
            <Text style={styles.totalLabel}>Total</Text>
            <Text style={styles.totalValue}>{fmt(p.total)}</Text>
          </View>
        </View>

        {/* Plans & Calculations (separate from building total) */}
        {(p.plans > 0 || p.calculations > 0) && (
          <View style={styles.section}>
            <Text style={styles.sectionTitle}>Plans &amp; Calculations</Text>
            <PriceLine label="Plans" value={p.plans} />
            <PriceLine label="Calculations" value={p.calculations} />
            {(p.plans > 0 || p.calculations > 0) && (
              <>
                <View style={styles.divider} />
                <View style={styles.row}>
                  <Text style={{ fontWeight: "bold" }}>Plans Total</Text>
                  <Text style={{ fontWeight: "bold" }}>{fmt(p.plans + p.calculations)}</Text>
                </View>
              </>
            )}
            <Text style={{ color: "#888", fontSize: 8, marginTop: 4 }}>
              Plans and calculations are billed separately and are not included in the building total above.
            </Text>
          </View>
        )}

        {/* Notes */}
        {quote.notes && (
          <View style={styles.section}>
            <Text style={styles.sectionTitle}>Notes</Text>
            <Text style={styles.label}>{quote.notes}</Text>
          </View>
        )}

        {/* Footer */}
        <Text style={styles.footer}>
          This quote is valid for 30 days from the date of issue. Prices subject to change.
        </Text>
      </Page>

      {/* Page 2: Snow Load Breakdown */}
      <SnowBreakdownPage quote={quote} />
    </Document>
  );
}

export async function generateQuotePdf(quote: Quote): Promise<Buffer> {
  return renderToBuffer(<QuoteDocument quote={quote} />) as Promise<Buffer>;
}
