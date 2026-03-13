"use client";

import type { PriceBreakdown } from "@/types/pricing";
import { formatCurrency } from "@/lib/utils";
import { Separator } from "@/components/ui/separator";
import { DEFAULT_DISCLAIMERS } from "@/lib/pricing/constants";

interface PriceSummaryProps {
  breakdown: PriceBreakdown | null;
  isWidespan: boolean;
  disclaimers?: string[];
}

function LineItem({ label, value }: { label: string; value: number }) {
  if (value === 0) return null;
  return (
    <div className="flex justify-between text-sm">
      <span className="text-muted-foreground">{label}</span>
      <span className="font-medium">{formatCurrency(value)}</span>
    </div>
  );
}

export function PriceSummary({ breakdown, isWidespan, disclaimers: disclaimersProp }: PriceSummaryProps) {
  const disclaimers = disclaimersProp ?? DEFAULT_DISCLAIMERS;
  if (!breakdown) {
    return (
      <div className="rounded-lg border bg-muted/30 p-6 text-center text-sm text-muted-foreground">
        Configure your building to see pricing.
      </div>
    );
  }

  return (
    <div className="space-y-4">
      {/* Main Price Breakdown */}
      <div className="space-y-3">
        <h3 className="font-semibold">Price Breakdown</h3>

        <div className="space-y-1.5">
          <LineItem label="Base Price" value={breakdown.basePrice} />
          {!isWidespan && <LineItem label="Roof Style" value={breakdown.roofStyle} />}
          <LineItem label="Leg Height" value={breakdown.legs} />
          <LineItem label="Sides" value={breakdown.sides} />
          <LineItem label="Ends" value={breakdown.ends} />
          <LineItem label="Walk-In Doors" value={breakdown.walkInDoors} />
          <LineItem label="Windows" value={breakdown.windows} />
          <LineItem label="Roll-Up Doors (Ends)" value={breakdown.rollUpDoorsEnds} />
          <LineItem label="Roll-Up Doors (Sides)" value={breakdown.rollUpDoorsSides} />
          <LineItem label="Insulation" value={breakdown.insulation} />
          {isWidespan && <LineItem label="Wainscot" value={breakdown.wainscot} />}
          {breakdown.contactEngineer ? (
            <div className="flex justify-between text-sm">
              <span className="text-muted-foreground">Snow/Wind Engineering</span>
              <span className="font-medium text-amber-500">Contact Engineer</span>
            </div>
          ) : (
            <LineItem label="Snow/Wind Engineering" value={breakdown.snowEngineering} />
          )}
          <LineItem label="Diagonal Bracing" value={breakdown.diagonalBracing} />
        </div>

        <Separator />

        <div className="flex justify-between text-sm font-medium">
          <span>Subtotal</span>
          <span>{formatCurrency(breakdown.subtotal)}</span>
        </div>

        <div className="flex justify-between text-sm">
          <span className="text-muted-foreground">
            Tax ({(breakdown.taxRate * 100).toFixed(2)}%)
          </span>
          <span>{formatCurrency(breakdown.taxAmount)}</span>
        </div>

        <LineItem label="Labor / Equipment" value={breakdown.laborEquipment} />

        <Separator />

        <div className="flex justify-between text-lg font-bold">
          <span>Total</span>
          <span>{formatCurrency(breakdown.total)}</span>
        </div>
        <p className="text-xs text-muted-foreground">
          Please add 15% for DROP OFF ONLY or 25% for INSTALL.
        </p>
      </div>

      {/* Additional Costs (separate from building price) */}
      {(breakdown.plans > 0 || breakdown.calculations > 0) && (
        <div className="rounded-lg border bg-muted/30 p-4 space-y-3">
          <h3 className="font-semibold text-sm">Additional Costs</h3>
          <div className="space-y-1.5">
            <LineItem label="Specific Plans Cost" value={breakdown.plans} />
            <LineItem label="Calculations Cost" value={breakdown.calculations} />
          </div>
          <Separator />
          <div className="flex justify-between text-sm font-medium">
            <span>Additional Total</span>
            <span>{formatCurrency(breakdown.plans + breakdown.calculations)}</span>
          </div>
          <p className="text-xs text-amber-500 font-medium">Not Included in Balance Due</p>
        </div>
      )}

      {/* Disclaimers */}
      <div className="space-y-2 pt-2">
        {disclaimers.map((text, i) => {
          const isSnowWarning = text.toLowerCase().includes("snow concerns");
          const isEngineeringWarning = text.toLowerCase().includes("contact the engineering");
          const isExcludes = text.toLowerCase().includes("quote excludes");
          let className = "text-[10px] text-muted-foreground leading-tight";
          if (isSnowWarning) className = "text-[10px] text-amber-500 leading-tight";
          else if (isEngineeringWarning) className = "text-[10px] text-destructive font-medium leading-tight";
          else if (isExcludes) className = "text-[10px] font-semibold leading-tight";
          return <p key={i} className={className}>{text}</p>;
        })}
      </div>
    </div>
  );
}
