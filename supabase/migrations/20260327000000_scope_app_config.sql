-- Scope asc_app_config by region_id and spreadsheet_type
-- so different regions and building sizes don't overwrite each other.

-- Add nullable scope columns (NULL = global default)
ALTER TABLE asc_app_config
  ADD COLUMN region_id UUID REFERENCES asc_regions(id) ON DELETE CASCADE,
  ADD COLUMN spreadsheet_type TEXT CHECK (spreadsheet_type IS NULL OR spreadsheet_type IN ('standard', 'widespan'));

-- Drop old PK (key alone) and add a surrogate PK
ALTER TABLE asc_app_config DROP CONSTRAINT asc_app_config_pkey;
ALTER TABLE asc_app_config ADD COLUMN id SERIAL PRIMARY KEY;

-- Unique constraint on (key, region_id, spreadsheet_type) with NULL handling
-- COALESCE ensures NULLs compare as equal for the unique check
CREATE UNIQUE INDEX asc_app_config_scope_uq
  ON asc_app_config (
    key,
    COALESCE(region_id, '00000000-0000-0000-0000-000000000000'::uuid),
    COALESCE(spreadsheet_type, '__global__')
  );

-- Index for scoped lookups
CREATE INDEX asc_app_config_region_idx ON asc_app_config (region_id, spreadsheet_type);

-- Function for atomic scoped upsert (DELETE + INSERT in one transaction)
-- Needed because Supabase JS client can't do ON CONFLICT with expression indexes
CREATE OR REPLACE FUNCTION upsert_scoped_config(
  p_key TEXT,
  p_value JSONB,
  p_region_id UUID,
  p_spreadsheet_type TEXT,
  p_updated_by TEXT
) RETURNS VOID AS $$
BEGIN
  DELETE FROM asc_app_config
    WHERE key = p_key
      AND region_id IS NOT DISTINCT FROM p_region_id
      AND spreadsheet_type IS NOT DISTINCT FROM p_spreadsheet_type;

  INSERT INTO asc_app_config (key, value, region_id, spreadsheet_type, updated_at, updated_by)
  VALUES (p_key, p_value, p_region_id, p_spreadsheet_type, now(), p_updated_by);
END;
$$ LANGUAGE plpgsql;
