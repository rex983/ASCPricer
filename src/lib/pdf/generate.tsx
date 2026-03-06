import React from "react";
import {
  Document,
  Page,
  Text,
  View,
  StyleSheet,
  renderToBuffer,
} from "@react-pdf/renderer";
import type { Quote } from "@/types/quote";
import type { PriceBreakdown } from "@/types/pricing";

const styles = StyleSheet.create({
  page: { padding: 40, fontSize: 10, fontFamily: "Helvetica" },
  header: { flexDirection: "row", justifyContent: "space-between", marginBottom: 20 },
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
});

function fmt(n: number): string {
  return "$" + n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function PriceLine({ label, value }: { label: string; value: number }) {
  if (value === 0) return null;
  return (
    <View style={styles.row}>
      <Text style={styles.label}>{label}</Text>
      <Text>{fmt(value)}</Text>
    </View>
  );
}

function QuoteDocument({ quote }: { quote: Quote }) {
  const p = quote.pricing as PriceBreakdown;
  const c = quote.config;

  return (
    <Document>
      <Page size="LETTER" style={styles.page}>
        {/* Header */}
        <View style={styles.header}>
          <View>
            <Text style={styles.title}>American Steel Carports</Text>
            <Text style={styles.subtitle}>Building Quote</Text>
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
          <PriceLine label="Base Price" value={p.basePrice} />
          <PriceLine label="Roof Style" value={p.roofStyle} />
          <PriceLine label="Leg Height" value={p.legs} />
          <PriceLine label="Sides" value={p.sides} />
          <PriceLine label="Ends" value={p.ends} />
          <PriceLine label="Walk-In Doors" value={p.walkInDoors} />
          <PriceLine label="Windows" value={p.windows} />
          <PriceLine label="Roll-Up Doors (Ends)" value={p.rollUpDoorsEnds} />
          <PriceLine label="Roll-Up Doors (Sides)" value={p.rollUpDoorsSides} />
          <PriceLine label="Insulation" value={p.insulation} />
          <PriceLine label="Wainscot" value={p.wainscot} />
          <PriceLine label="Snow/Wind Engineering" value={p.snowEngineering} />
          <PriceLine label="Diagonal Bracing" value={p.diagonalBracing} />
          <PriceLine label="Plans" value={p.plans} />

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
    </Document>
  );
}

export async function generateQuotePdf(quote: Quote): Promise<Buffer> {
  return renderToBuffer(<QuoteDocument quote={quote} />) as Promise<Buffer>;
}
