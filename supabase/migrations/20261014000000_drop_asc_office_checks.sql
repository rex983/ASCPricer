-- Offices are managed in the BBD Launcher (launcher_offices); admins can add
-- new ones at any time. Drop the old hardcoded office CHECK constraints
-- (office IN ('Harbor','BST','RnD')) on ASC-owned tables so any launcher
-- office can be stored. Validation happens in the app against
-- launcher_offices. profiles.office belongs to the launcher and is not
-- touched here. Idempotent.

DO $$
DECLARE
  r record;
BEGIN
  FOR r IN
    SELECT c.conrelid::regclass AS tbl, c.conname
    FROM pg_constraint c
    WHERE c.contype = 'c'
      AND c.conrelid IN (
        SELECT oid FROM pg_class
        WHERE relname IN ('asc_customers', 'asc_quotes', 'asc_sales_reps')
          AND relnamespace = 'public'::regnamespace
      )
      AND pg_get_constraintdef(c.oid) ILIKE '%office%'
  LOOP
    EXECUTE format('ALTER TABLE %s DROP CONSTRAINT IF EXISTS %I', r.tbl, r.conname);
  END LOOP;
END $$;
