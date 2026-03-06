"use client";

import { useMemo } from "react";
import type { BuildingConfig, PriceBreakdown, PricingMatrices } from "@/types/pricing";
import { calculatePrice } from "@/lib/pricing/engine";

/**
 * React hook for live price calculation.
 * Recalculates whenever config or matrices change.
 */
export function usePricingEngine(
  config: BuildingConfig | null,
  matrices: PricingMatrices | null
): PriceBreakdown | null {
  return useMemo(() => {
    if (!config || !matrices) return null;

    try {
      return calculatePrice(config, matrices);
    } catch (error) {
      console.error("Pricing engine error:", error);
      return null;
    }
  }, [config, matrices]);
}
