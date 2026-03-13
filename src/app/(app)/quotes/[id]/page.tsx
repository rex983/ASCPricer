"use client";

import { useEffect, useState } from "react";
import { useParams, useRouter } from "next/navigation";
import { AppHeader } from "@/components/layout/app-header";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Separator } from "@/components/ui/separator";
import { formatCurrency, formatDate } from "@/lib/utils";
import type { Quote } from "@/types/quote";
import type { PriceBreakdown } from "@/types/pricing";
import { ArrowLeft, FileDown, Loader2, Save } from "lucide-react";

function PriceLine({ label, value }: { label: string; value: number }) {
  if (value === 0) return null;
  return (
    <div className="flex justify-between text-sm">
      <span className="text-muted-foreground">{label}</span>
      <span>{formatCurrency(value)}</span>
    </div>
  );
}

export default function QuoteDetailPage() {
  const { id } = useParams<{ id: string }>();
  const router = useRouter();
  const [quote, setQuote] = useState<Quote | null>(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [editFields, setEditFields] = useState({
    status: "",
    customer_name: "",
    customer_email: "",
    customer_phone: "",
    customer_address: "",
    customer_city: "",
    customer_state: "",
    customer_zip: "",
    notes: "",
  });

  useEffect(() => {
    fetch(`/api/quotes/${id}`)
      .then((r) => r.json())
      .then((data) => {
        if (data.id) {
          setQuote(data);
          setEditFields({
            status: data.status,
            customer_name: data.customer_name || "",
            customer_email: data.customer_email || "",
            customer_phone: data.customer_phone || "",
            customer_address: data.customer_address || "",
            customer_city: data.customer_city || "",
            customer_state: data.customer_state || "",
            customer_zip: data.customer_zip || "",
            notes: data.notes || "",
          });
        }
      })
      .catch(() => {})
      .finally(() => setLoading(false));
  }, [id]);

  const handleSave = async () => {
    setSaving(true);
    try {
      const res = await fetch(`/api/quotes/${id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(editFields),
      });
      const data = await res.json();
      if (data.id) setQuote(data);
    } catch {
      // ignore
    } finally {
      setSaving(false);
    }
  };

  if (loading) {
    return (
      <>
        <AppHeader title="Quote" />
        <div className="flex-1 flex items-center justify-center">
          <Loader2 className="h-8 w-8 animate-spin text-muted-foreground" />
        </div>
      </>
    );
  }

  if (!quote) {
    return (
      <>
        <AppHeader title="Quote Not Found" />
        <div className="flex-1 p-6 text-center text-muted-foreground">
          <p>This quote could not be found.</p>
          <Button variant="outline" className="mt-4" onClick={() => router.push("/quotes")}>
            Back to Quotes
          </Button>
        </div>
      </>
    );
  }

  const p = quote.pricing as PriceBreakdown;

  return (
    <>
      <AppHeader title={quote.quote_number} />
      <div className="flex-1 p-6">
        <div className="mx-auto max-w-5xl space-y-6">
          {/* Header bar */}
          <div className="flex items-center gap-4">
            <Button variant="outline" size="sm" onClick={() => router.push("/quotes")}>
              <ArrowLeft className="mr-1 h-4 w-4" /> Back
            </Button>
            <h2 className="text-xl font-semibold">{quote.quote_number}</h2>
            <Badge>{quote.status}</Badge>
            <div className="ml-auto flex items-center gap-3">
              <span className="text-sm text-muted-foreground">
                Created {formatDate(quote.created_at)}
              </span>
              <Button variant="outline" size="sm" asChild>
                <a href={`/api/quotes/${id}/pdf`} target="_blank" rel="noopener noreferrer">
                  <FileDown className="mr-1 h-4 w-4" /> PDF
                </a>
              </Button>
            </div>
          </div>

          <div className="grid gap-6 lg:grid-cols-[1fr_320px]">
            {/* Left: Customer + Config */}
            <div className="space-y-6">
              {/* Customer Info */}
              <Card>
                <CardHeader className="pb-4">
                  <CardTitle className="text-base">Customer Information</CardTitle>
                </CardHeader>
                <CardContent className="space-y-3">
                  <div className="grid grid-cols-2 gap-3">
                    <div className="space-y-1">
                      <Label className="text-xs">Name</Label>
                      <Input
                        value={editFields.customer_name}
                        onChange={(e) => setEditFields((f) => ({ ...f, customer_name: e.target.value }))}
                      />
                    </div>
                    <div className="space-y-1">
                      <Label className="text-xs">Email</Label>
                      <Input
                        value={editFields.customer_email}
                        onChange={(e) => setEditFields((f) => ({ ...f, customer_email: e.target.value }))}
                      />
                    </div>
                    <div className="space-y-1">
                      <Label className="text-xs">Phone</Label>
                      <Input
                        value={editFields.customer_phone}
                        onChange={(e) => setEditFields((f) => ({ ...f, customer_phone: e.target.value }))}
                      />
                    </div>
                    <div className="space-y-1">
                      <Label className="text-xs">Address</Label>
                      <Input
                        value={editFields.customer_address}
                        onChange={(e) => setEditFields((f) => ({ ...f, customer_address: e.target.value }))}
                      />
                    </div>
                    <div className="space-y-1">
                      <Label className="text-xs">City</Label>
                      <Input
                        value={editFields.customer_city}
                        onChange={(e) => setEditFields((f) => ({ ...f, customer_city: e.target.value }))}
                      />
                    </div>
                    <div className="grid grid-cols-2 gap-3">
                      <div className="space-y-1">
                        <Label className="text-xs">State</Label>
                        <Input
                          value={editFields.customer_state}
                          onChange={(e) => setEditFields((f) => ({ ...f, customer_state: e.target.value }))}
                        />
                      </div>
                      <div className="space-y-1">
                        <Label className="text-xs">ZIP</Label>
                        <Input
                          value={editFields.customer_zip}
                          onChange={(e) => setEditFields((f) => ({ ...f, customer_zip: e.target.value }))}
                        />
                      </div>
                    </div>
                  </div>
                  <div className="space-y-1">
                    <Label className="text-xs">Notes</Label>
                    <textarea
                      className="flex min-h-[80px] w-full rounded-md border border-input bg-background px-3 py-2 text-sm"
                      value={editFields.notes}
                      onChange={(e) => setEditFields((f) => ({ ...f, notes: e.target.value }))}
                    />
                  </div>
                </CardContent>
              </Card>

              {/* Status + Save */}
              <Card>
                <CardContent className="pt-6">
                  <div className="flex items-end gap-4">
                    <div className="space-y-1">
                      <Label className="text-xs">Status</Label>
                      <Select
                        value={editFields.status}
                        onValueChange={(v) => setEditFields((f) => ({ ...f, status: v }))}
                      >
                        <SelectTrigger className="w-36">
                          <SelectValue />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="draft">Draft</SelectItem>
                          <SelectItem value="sent">Sent</SelectItem>
                          <SelectItem value="accepted">Accepted</SelectItem>
                          <SelectItem value="expired">Expired</SelectItem>
                        </SelectContent>
                      </Select>
                    </div>
                    <Button onClick={handleSave} disabled={saving}>
                      {saving ? <Loader2 className="mr-1 h-4 w-4 animate-spin" /> : <Save className="mr-1 h-4 w-4" />}
                      Save Changes
                    </Button>
                  </div>
                </CardContent>
              </Card>

              {/* Building Config Summary */}
              <Card>
                <CardHeader className="pb-4">
                  <CardTitle className="text-base">Building Configuration</CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="grid grid-cols-2 gap-x-6 gap-y-2 text-sm sm:grid-cols-4">
                    <div><span className="text-muted-foreground">Width:</span> {quote.config.width}&apos;</div>
                    <div><span className="text-muted-foreground">Length:</span> {quote.config.length}&apos;</div>
                    <div><span className="text-muted-foreground">Height:</span> {quote.config.height}&apos;</div>
                    <div><span className="text-muted-foreground">Gauge:</span> {quote.config.gauge}G</div>
                    <div><span className="text-muted-foreground">Roof:</span> {quote.config.roofStyle}</div>
                    <div><span className="text-muted-foreground">Sides:</span> {quote.config.sidesCoverage} x{quote.config.sidesQty}</div>
                    <div><span className="text-muted-foreground">Ends:</span> {quote.config.endType} x{quote.config.endsQty}</div>
                    <div><span className="text-muted-foreground">Insulation:</span> {quote.config.insulationType}</div>
                  </div>
                </CardContent>
              </Card>
            </div>

            {/* Right: Price Breakdown */}
            <div className="lg:sticky lg:top-6 lg:self-start">
              <Card>
                <CardHeader className="pb-4">
                  <CardTitle className="text-base">Price Breakdown</CardTitle>
                </CardHeader>
                <CardContent className="space-y-2">
                  <PriceLine label="Base Price" value={p.basePrice} />
                  <PriceLine label="Roof Style" value={p.roofStyle} />
                  <PriceLine label="Leg Height" value={p.legs} />
                  <PriceLine label="Sides" value={p.sides} />
                  <PriceLine label="Ends" value={p.ends} />
                  <PriceLine label="Walk-In Doors" value={p.walkInDoors} />
                  <PriceLine label="Windows" value={p.windows} />
                  <PriceLine label="Roll-Up (Ends)" value={p.rollUpDoorsEnds} />
                  <PriceLine label="Roll-Up (Sides)" value={p.rollUpDoorsSides} />
                  <PriceLine label="Insulation" value={p.insulation} />
                  <PriceLine label="Wainscot" value={p.wainscot} />
                  <PriceLine label="Snow/Wind" value={p.snowEngineering} />
                  <PriceLine label="Diagonal Bracing" value={p.diagonalBracing} />

                  <Separator />

                  <div className="flex justify-between text-sm font-medium">
                    <span>Subtotal</span>
                    <span>{formatCurrency(p.subtotal)}</span>
                  </div>
                  <div className="flex justify-between text-sm">
                    <span className="text-muted-foreground">Tax ({(p.taxRate * 100).toFixed(2)}%)</span>
                    <span>{formatCurrency(p.taxAmount)}</span>
                  </div>
                  <PriceLine label="Labor / Equipment" value={p.laborEquipment} />

                  <Separator />

                  <div className="flex justify-between text-lg font-bold">
                    <span>Total</span>
                    <span>{formatCurrency(p.total)}</span>
                  </div>
                </CardContent>
              </Card>

              {(p.plans > 0 || p.calculations > 0) && (
                <Card className="mt-4 bg-muted/30">
                  <CardHeader className="pb-4">
                    <CardTitle className="text-base">Additional Costs</CardTitle>
                  </CardHeader>
                  <CardContent className="space-y-2">
                    <PriceLine label="Specific Plans Cost" value={p.plans} />
                    <PriceLine label="Calculations Cost" value={p.calculations ?? 0} />
                    <Separator />
                    <div className="flex justify-between text-sm font-medium">
                      <span>Additional Total</span>
                      <span>{formatCurrency(p.plans + (p.calculations ?? 0))}</span>
                    </div>
                  </CardContent>
                </Card>
              )}
            </div>
          </div>
        </div>
      </div>
    </>
  );
}
