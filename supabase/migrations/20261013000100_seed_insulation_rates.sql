-- Move the insulation $/sqft rates out of code and into asc_app_config so
-- admins can edit them on the Variables page.
--
-- Values match src/lib/pricing/constants.ts (FIBERGLASS_RATE = 2.25,
-- THERMAL_RATE = 1.65), which remain the code fallback, so pricing output is
-- unchanged. Global row only (region_id / spreadsheet_type NULL).
--
-- Idempotent: inserts only if the global row doesn't exist yet, and never
-- overwrites an edited value.

INSERT INTO asc_app_config (key, value)
SELECT 'insulation_rates', '{"fiberglass": 2.25, "thermal": 1.65}'::jsonb
WHERE NOT EXISTS (
  SELECT 1 FROM asc_app_config
  WHERE key = 'insulation_rates'
    AND region_id IS NULL
    AND spreadsheet_type IS NULL
);
