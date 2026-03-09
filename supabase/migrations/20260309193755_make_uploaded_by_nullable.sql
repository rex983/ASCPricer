-- Make uploaded_by nullable so credentials-based admin login works
-- (credentials login uses "admin-001" as profileId which isn't a valid UUID)
ALTER TABLE asc_uploads ALTER COLUMN uploaded_by DROP NOT NULL;
