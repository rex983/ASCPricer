import type { BuildingConfig, StandardMatrices, WidespanMatrices } from "@/types/pricing";
import { STANDARD_WIDTHS, WIDESPAN_WIDTHS, ROOF_STYLE_KEYS } from "./constants";
import { nearestBucket, roundToIncrement } from "./lookups";

/**
 * Resolve user inputs into lookup keys for the standard pricing engine.
 * This replicates the "Pricing - Changers" sheet logic.
 */
export function resolveStandardKeys(config: BuildingConfig, matrices: StandardMatrices) {
  const width = nearestBucket(config.width, STANDARD_WIDTHS);
  const length = roundToIncrement(config.length, 5);
  const gauge = config.gauge;
  const roofKey = ROOF_STYLE_KEYS[config.roofStyle];

  return {
    width,
    length,
    gauge,
    basePriceKey: `${width}-${gauge}G`, // e.g., "24-14G"
    roofStyleKey: `${roofKey}-${width}`, // e.g., "AFV-24"
    roofKey,
    sidesOrientationKey: config.sidesOrientation === "vertical" ? "V" : "HZ",
    endsOrientationKey: config.endsOrientation === "vertical" ? "V" : "HZ",
  };
}

/**
 * Resolve user inputs into lookup keys for the widespan pricing engine.
 * This replicates the widespan "Changers" sheet logic.
 */
export function resolveWidespanKeys(config: BuildingConfig, matrices: WidespanMatrices) {
  const width = nearestBucket(config.width, WIDESPAN_WIDTHS);
  const length = roundToIncrement(config.length, 5);
  const gauge = config.gauge;

  return {
    width,
    length,
    gauge,
    basePriceKey: `${width}-${gauge}G`,
  };
}
