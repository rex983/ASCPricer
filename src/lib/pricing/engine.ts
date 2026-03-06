import type { BuildingConfig, PriceBreakdown, PricingMatrices } from "@/types/pricing";
import { calculateStandardPrice } from "./engine-standard";
import { calculateWidespanPrice } from "./engine-widespan";
import { getSpreadsheetType } from "./building-type";

/**
 * Main pricing engine orchestrator.
 * Routes to standard or widespan engine based on building width.
 *
 * This function is isomorphic — it runs in both browser and server.
 * The server re-validates when saving quotes.
 */
export function calculatePrice(
  config: BuildingConfig,
  matrices: PricingMatrices
): PriceBreakdown {
  const type = getSpreadsheetType(config.width);

  if (type === "standard" && matrices.type === "standard") {
    return calculateStandardPrice(config, matrices);
  }

  if (type === "widespan" && matrices.type === "widespan") {
    return calculateWidespanPrice(config, matrices);
  }

  throw new Error(
    `Mismatched building width (${config.width}) and pricing data type (${matrices.type}). ` +
    `Width ≤30 requires standard matrices, width 32+ requires widespan matrices.`
  );
}
