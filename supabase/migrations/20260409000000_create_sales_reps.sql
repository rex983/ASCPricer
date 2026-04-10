-- Sales reps table
CREATE TABLE IF NOT EXISTS asc_sales_reps (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  profile_id UUID REFERENCES profiles(id) ON DELETE SET NULL,
  name TEXT NOT NULL,
  email TEXT NOT NULL,
  phone TEXT,
  office TEXT NOT NULL CHECK (office IN ('Harbor', 'Marion')),
  territory TEXT,
  commission_rate NUMERIC(5,2) DEFAULT 0,
  is_active BOOLEAN NOT NULL DEFAULT true,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Auto-update updated_at
CREATE TRIGGER set_asc_sales_reps_updated_at
  BEFORE UPDATE ON asc_sales_reps
  FOR EACH ROW
  EXECUTE FUNCTION update_updated_at();

-- Enable RLS
ALTER TABLE asc_sales_reps ENABLE ROW LEVEL SECURITY;

-- Service role has full access
CREATE POLICY "service_role_all_sales_reps" ON asc_sales_reps
  FOR ALL TO service_role USING (true) WITH CHECK (true);

-- Authenticated users can read sales reps
CREATE POLICY "authenticated_read_sales_reps" ON asc_sales_reps
  FOR SELECT TO authenticated USING (true);

-- Add assigned sales rep to customers
ALTER TABLE asc_customers ADD COLUMN IF NOT EXISTS assigned_rep_id UUID REFERENCES asc_sales_reps(id) ON DELETE SET NULL;

-- Indexes
CREATE INDEX IF NOT EXISTS idx_asc_sales_reps_office ON asc_sales_reps(office);
CREATE INDEX IF NOT EXISTS idx_asc_sales_reps_profile_id ON asc_sales_reps(profile_id);
CREATE INDEX IF NOT EXISTS idx_asc_sales_reps_is_active ON asc_sales_reps(is_active);
CREATE INDEX IF NOT EXISTS idx_asc_customers_assigned_rep ON asc_customers(assigned_rep_id);
