-- Add office column to profiles (shared table)
ALTER TABLE profiles ADD COLUMN IF NOT EXISTS office TEXT CHECK (office IN ('Harbor', 'Marion'));

-- Create customers table
CREATE TABLE IF NOT EXISTS asc_customers (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name TEXT NOT NULL,
  email TEXT,
  phone TEXT,
  address TEXT,
  city TEXT,
  state TEXT,
  zip TEXT,
  notes TEXT,
  office TEXT NOT NULL CHECK (office IN ('Harbor', 'Marion')),
  created_by UUID REFERENCES profiles(id),
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Auto-update updated_at on customers
CREATE TRIGGER set_asc_customers_updated_at
  BEFORE UPDATE ON asc_customers
  FOR EACH ROW
  EXECUTE FUNCTION update_updated_at();

-- Add customer_id and office to quotes
ALTER TABLE asc_quotes ADD COLUMN IF NOT EXISTS customer_id UUID REFERENCES asc_customers(id);
ALTER TABLE asc_quotes ADD COLUMN IF NOT EXISTS office TEXT CHECK (office IN ('Harbor', 'Marion'));

-- Enable RLS on customers
ALTER TABLE asc_customers ENABLE ROW LEVEL SECURITY;

-- Service role has full access
CREATE POLICY "service_role_all_customers" ON asc_customers
  FOR ALL TO service_role USING (true) WITH CHECK (true);

-- Authenticated users can read customers (API handles role-based filtering)
CREATE POLICY "authenticated_read_customers" ON asc_customers
  FOR SELECT TO authenticated USING (true);

-- Index for common lookups
CREATE INDEX IF NOT EXISTS idx_asc_customers_office ON asc_customers(office);
CREATE INDEX IF NOT EXISTS idx_asc_customers_created_by ON asc_customers(created_by);
CREATE INDEX IF NOT EXISTS idx_asc_quotes_office ON asc_quotes(office);
CREATE INDEX IF NOT EXISTS idx_asc_quotes_created_by ON asc_quotes(created_by);
CREATE INDEX IF NOT EXISTS idx_asc_customers_name ON asc_customers(name);
