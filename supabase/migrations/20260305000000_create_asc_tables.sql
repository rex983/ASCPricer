-- American Steel Carports Pricing Tables
-- These tables live in the shared Supabase instance alongside bbd-launcher tables

-- Regions (e.g., "AZ/CO/UT", "OH/PA/NY")
CREATE TABLE asc_regions (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name TEXT NOT NULL,
  slug TEXT NOT NULL UNIQUE,
  states TEXT[] NOT NULL DEFAULT '{}',
  is_active BOOLEAN NOT NULL DEFAULT true,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Upload tracking
CREATE TABLE asc_uploads (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  region_id UUID NOT NULL REFERENCES asc_regions(id) ON DELETE CASCADE,
  uploaded_by UUID REFERENCES profiles(id),
  filename TEXT NOT NULL,
  spreadsheet_type TEXT NOT NULL CHECK (spreadsheet_type IN ('standard', 'widespan')),
  sheet_count INT,
  status TEXT NOT NULL DEFAULT 'processing' CHECK (status IN ('processing', 'success', 'failed')),
  error_message TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Parsed pricing data (JSON blobs)
CREATE TABLE asc_pricing_data (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  region_id UUID NOT NULL REFERENCES asc_regions(id) ON DELETE CASCADE,
  version INT NOT NULL DEFAULT 1,
  is_current BOOLEAN NOT NULL DEFAULT false,
  spreadsheet_type TEXT NOT NULL CHECK (spreadsheet_type IN ('standard', 'widespan')),
  matrices JSONB NOT NULL DEFAULT '{}',
  upload_id UUID REFERENCES asc_uploads(id),
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Only one current pricing data per region+type
CREATE UNIQUE INDEX idx_asc_pricing_data_current
  ON asc_pricing_data (region_id, spreadsheet_type)
  WHERE is_current = true;

-- Quotes
CREATE TABLE asc_quotes (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  quote_number TEXT NOT NULL UNIQUE,
  region_id UUID NOT NULL REFERENCES asc_regions(id),
  pricing_data_id UUID REFERENCES asc_pricing_data(id),
  created_by UUID REFERENCES profiles(id),
  status TEXT NOT NULL DEFAULT 'draft' CHECK (status IN ('draft', 'sent', 'accepted', 'expired')),
  customer_name TEXT,
  customer_email TEXT,
  customer_phone TEXT,
  customer_address TEXT,
  customer_city TEXT,
  customer_state TEXT,
  customer_zip TEXT,
  config JSONB NOT NULL DEFAULT '{}',
  pricing JSONB NOT NULL DEFAULT '{}',
  subtotal NUMERIC(12,2) NOT NULL DEFAULT 0,
  tax_rate NUMERIC(5,4) NOT NULL DEFAULT 0,
  tax_amount NUMERIC(12,2) NOT NULL DEFAULT 0,
  total NUMERIC(12,2) NOT NULL DEFAULT 0,
  notes TEXT,
  valid_until DATE,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Quote number sequence
CREATE TABLE asc_quote_sequence (
  year INT PRIMARY KEY,
  last_number INT NOT NULL DEFAULT 0
);

-- Function to generate next quote number (ASC-2026-0001)
CREATE OR REPLACE FUNCTION next_quote_number()
RETURNS TEXT
LANGUAGE plpgsql
AS $$
DECLARE
  current_year INT := EXTRACT(YEAR FROM now());
  next_num INT;
BEGIN
  INSERT INTO asc_quote_sequence (year, last_number)
  VALUES (current_year, 1)
  ON CONFLICT (year)
  DO UPDATE SET last_number = asc_quote_sequence.last_number + 1
  RETURNING last_number INTO next_num;

  RETURN 'ASC-' || current_year || '-' || LPAD(next_num::TEXT, 4, '0');
END;
$$;

-- Updated_at triggers
CREATE OR REPLACE FUNCTION update_updated_at()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = now();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trg_asc_regions_updated_at
  BEFORE UPDATE ON asc_regions
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();

CREATE TRIGGER trg_asc_quotes_updated_at
  BEFORE UPDATE ON asc_quotes
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();

-- Indexes
CREATE INDEX idx_asc_pricing_data_region ON asc_pricing_data(region_id);
CREATE INDEX idx_asc_quotes_region ON asc_quotes(region_id);
CREATE INDEX idx_asc_quotes_created_by ON asc_quotes(created_by);
CREATE INDEX idx_asc_quotes_status ON asc_quotes(status);
CREATE INDEX idx_asc_uploads_region ON asc_uploads(region_id);

-- RLS policies (basic — will be refined in Phase 7)
ALTER TABLE asc_regions ENABLE ROW LEVEL SECURITY;
ALTER TABLE asc_pricing_data ENABLE ROW LEVEL SECURITY;
ALTER TABLE asc_uploads ENABLE ROW LEVEL SECURITY;
ALTER TABLE asc_quotes ENABLE ROW LEVEL SECURITY;

-- Allow authenticated read on regions and pricing data
CREATE POLICY "Authenticated users can read regions"
  ON asc_regions FOR SELECT TO authenticated USING (true);

CREATE POLICY "Authenticated users can read pricing data"
  ON asc_pricing_data FOR SELECT TO authenticated USING (true);

CREATE POLICY "Authenticated users can read uploads"
  ON asc_uploads FOR SELECT TO authenticated USING (true);

CREATE POLICY "Authenticated users can read quotes"
  ON asc_quotes FOR SELECT TO authenticated USING (true);

-- Service role has full access (admin operations go through service role)
CREATE POLICY "Service role full access regions"
  ON asc_regions FOR ALL TO service_role USING (true) WITH CHECK (true);

CREATE POLICY "Service role full access pricing_data"
  ON asc_pricing_data FOR ALL TO service_role USING (true) WITH CHECK (true);

CREATE POLICY "Service role full access uploads"
  ON asc_uploads FOR ALL TO service_role USING (true) WITH CHECK (true);

CREATE POLICY "Service role full access quotes"
  ON asc_quotes FOR ALL TO service_role USING (true) WITH CHECK (true);
