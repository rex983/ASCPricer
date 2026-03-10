-- Drop the FK constraint on uploaded_by entirely.
-- The credentials login uses a non-UUID profileId ("admin-001"),
-- and even with the column nullable, the Supabase JS client may
-- still send null which triggers the FK check.
ALTER TABLE asc_uploads DROP CONSTRAINT IF EXISTS asc_uploads_uploaded_by_fkey;
