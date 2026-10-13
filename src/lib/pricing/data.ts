import { revalidateTag, unstable_cache } from "next/cache";
import { createAdminClient } from "@/lib/supabase/admin";
import { FIBERGLASS_RATE, THERMAL_RATE } from "./constants";
import type { PricingMatrices } from "@/types/pricing";

/**
 * Server-side cached reads of pricing data. Everything here is shared by
 * all users, so it's cached in the Next data cache and invalidated by tag
 * from the admin routes that change it (upload, regions, config).
 *
 * Backstop TTL in case the tables are edited outside the app.
 */
const BACKSTOP_TTL_SECONDS = 300;

export const PRICING_DATA_TAG = "asc-pricing-data";
export const REGIONS_TAG = "asc-regions";
export const APP_CONFIG_TAG = "asc-app-config";

/** Expire the given cache tags immediately (route handlers can't use updateTag). */
export function invalidatePricingCache(...tags: string[]) {
  for (const tag of tags) revalidateTag(tag, { expire: 0 });
}

/** Error carrying the PostgREST/Postgres code so callers keep their status mapping. */
export class PricingDataError extends Error {
  constructor(message: string, public code?: string) {
    super(message);
  }
}

// ── Pricing data ─────────────────────────────────────────────

export interface PricingDataRow {
  id: string;
  region_id: string;
  spreadsheet_type: string;
  matrices: PricingMatrices;
  version: number;
  created_at: string;
}

/**
 * Current (or a specific historical) pricing row for a region.
 * Returns null when none exists. Throws PricingDataError on DB errors
 * (errors are never cached).
 */
export const getPricingRow = unstable_cache(
  async (regionId: string, versionId: string | null): Promise<PricingDataRow | null> => {
    const supabase = createAdminClient();
    let query = supabase
      .from("asc_pricing_data")
      .select("id, region_id, spreadsheet_type, matrices, version, created_at")
      .eq("region_id", regionId);
    query = versionId ? query.eq("id", versionId) : query.eq("is_current", true);

    const { data, error } = await query.single();
    if (error) {
      if (error.code === "PGRST116") return null;
      throw new PricingDataError(error.message, error.code);
    }
    return data as PricingDataRow;
  },
  ["asc-pricing-row"],
  { tags: [PRICING_DATA_TAG], revalidate: BACKSTOP_TTL_SECONDS }
);

export interface PricingVersion {
  id: string;
  version: number;
  spreadsheetType: string;
  isCurrent: boolean;
  createdAt: string;
  filename: string | null;
}

export const getPricingVersions = unstable_cache(
  async (regionId: string): Promise<PricingVersion[]> => {
    const supabase = createAdminClient();
    const { data, error } = await supabase
      .from("asc_pricing_data")
      .select(
        "id, version, spreadsheet_type, is_current, created_at, upload_id, asc_uploads(filename)"
      )
      .eq("region_id", regionId)
      .order("version", { ascending: false });

    if (error) throw new PricingDataError(error.message, error.code);

    // Flatten the upload filename into each row
    return (data || []).map((row) => {
      const upload = row.asc_uploads as unknown as { filename: string } | null;
      return {
        id: row.id,
        version: row.version,
        spreadsheetType: row.spreadsheet_type,
        isCurrent: row.is_current,
        createdAt: row.created_at,
        filename: upload?.filename || null,
      };
    });
  },
  ["asc-pricing-versions"],
  { tags: [PRICING_DATA_TAG], revalidate: BACKSTOP_TTL_SECONDS }
);

// ── Regions ──────────────────────────────────────────────────

export const getActiveRegions = unstable_cache(
  async (type: string | null) => {
    const supabase = createAdminClient();
    let query = supabase
      .from("asc_regions")
      .select("*")
      .eq("is_active", true)
      .order("name");
    if (type) query = query.eq("spreadsheet_type", type);

    const { data, error } = await query;
    if (error) throw new PricingDataError(error.message, error.code);
    return data;
  },
  ["asc-active-regions"],
  { tags: [REGIONS_TAG], revalidate: BACKSTOP_TTL_SECONDS }
);

// ── App config ───────────────────────────────────────────────

/**
 * Global config with optional region + spreadsheet-type overlay.
 * Returns {} if asc_app_config doesn't exist yet.
 */
export const getAppConfig = unstable_cache(
  async (
    regionId: string | null,
    spreadsheetType: string | null
  ): Promise<Record<string, unknown>> => {
    const supabase = createAdminClient();

    // 1. Always load global config (region_id IS NULL)
    const { data: globalRows, error } = await supabase
      .from("asc_app_config")
      .select("key, value")
      .is("region_id", null);

    if (error) {
      if (error.code === "42P01" || error.message?.includes("does not exist")) {
        return {};
      }
      throw new PricingDataError(error.message, error.code);
    }

    const config: Record<string, unknown> = {};
    for (const row of globalRows ?? []) {
      config[row.key] = row.value;
    }

    // 2. If region + type provided, overlay scoped config (takes priority)
    if (regionId && spreadsheetType) {
      const { data: scopedRows } = await supabase
        .from("asc_app_config")
        .select("key, value")
        .eq("region_id", regionId)
        .eq("spreadsheet_type", spreadsheetType);

      for (const row of scopedRows ?? []) {
        config[row.key] = row.value;
      }
    }

    return config;
  },
  ["asc-app-config"],
  { tags: [APP_CONFIG_TAG], revalidate: BACKSTOP_TTL_SECONDS }
);

// ── Insulation rates ─────────────────────────────────────────

export const INSULATION_RATES_KEY = "insulation_rates";

export interface InsulationRates {
  fiberglass: number;
  thermal: number;
}

export const DEFAULT_INSULATION_RATES: InsulationRates = {
  fiberglass: FIBERGLASS_RATE,
  thermal: THERMAL_RATE,
};

function isValidRate(v: unknown): v is number {
  return typeof v === "number" && Number.isFinite(v) && v > 0;
}

/** Validate an `insulation_rates` config value. */
export function parseInsulationRates(value: unknown): InsulationRates | null {
  if (typeof value !== "object" || value === null) return null;
  const { fiberglass, thermal } = value as Record<string, unknown>;
  if (!isValidRate(fiberglass) || !isValidRate(thermal)) return null;
  return { fiberglass, thermal };
}

/**
 * Per-sqft insulation rates from asc_app_config (global row), falling back
 * to the FIBERGLASS_RATE / THERMAL_RATE constants if the row is missing,
 * malformed, or the DB read fails.
 */
export async function getInsulationRates(): Promise<InsulationRates> {
  try {
    const config = await getAppConfig(null, null);
    return parseInsulationRates(config[INSULATION_RATES_KEY]) ?? DEFAULT_INSULATION_RATES;
  } catch (err) {
    console.error("Failed to load insulation rates, using defaults:", err);
    return DEFAULT_INSULATION_RATES;
  }
}

/**
 * Standard spreadsheets carry no insulation rates (the parser always stored
 * the constants), so the configured rates apply to standard matrices at
 * read time. Widespan spreadsheets have their own rates in the
 * "Insulation - Wainscot" sheet, which take precedence; the configured
 * rates are only their parse-time fallback.
 */
export function applyInsulationRates(
  matrices: PricingMatrices,
  rates: InsulationRates
): PricingMatrices {
  if (matrices.type !== "standard") return matrices;
  return {
    ...matrices,
    insulation: { fiberglassRate: rates.fiberglass, thermalRate: rates.thermal },
  };
}
