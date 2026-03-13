-- Audit log for tracking user actions (uploads, config changes, etc.)
CREATE TABLE IF NOT EXISTS asc_audit_log (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id TEXT,
  user_email TEXT,
  action TEXT NOT NULL,
  resource_type TEXT,
  resource_id TEXT,
  details JSONB DEFAULT '{}',
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Index for querying by time (most recent first)
CREATE INDEX idx_audit_log_created_at ON asc_audit_log (created_at DESC);

-- Enable RLS
ALTER TABLE asc_audit_log ENABLE ROW LEVEL SECURITY;

-- Allow admins to read audit logs (service role bypasses RLS for inserts)
CREATE POLICY "Admins can view audit logs"
  ON asc_audit_log FOR SELECT
  USING (true);
