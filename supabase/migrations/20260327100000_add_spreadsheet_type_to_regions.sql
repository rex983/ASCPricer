-- Add spreadsheet_type to asc_regions (may already exist from manual SQL)
-- Each row is a region+size combo (e.g. "AZ Standard" and "AZ Widespan" are separate)
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_name = 'asc_regions' AND column_name = 'spreadsheet_type'
  ) THEN
    ALTER TABLE asc_regions
      ADD COLUMN spreadsheet_type TEXT NOT NULL DEFAULT 'standard'
      CHECK (spreadsheet_type IN ('standard', 'widespan'));
  END IF;
END $$;

-- Replace the old unique constraint on slug with one on (slug, spreadsheet_type)
-- so the same state group can have both standard and widespan entries
DROP INDEX IF EXISTS asc_regions_slug_key;
CREATE UNIQUE INDEX IF NOT EXISTS asc_regions_slug_type_uq
  ON asc_regions (slug, spreadsheet_type);
