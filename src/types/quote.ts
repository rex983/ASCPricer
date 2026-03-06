import type { BuildingConfig, PriceBreakdown } from "./pricing";

export type QuoteStatus = "draft" | "sent" | "accepted" | "expired";

export interface Quote {
  id: string;
  quote_number: string;
  region_id: string;
  pricing_data_id: string | null;
  created_by: string | null;
  status: QuoteStatus;
  customer_name: string | null;
  customer_email: string | null;
  customer_phone: string | null;
  customer_address: string | null;
  customer_city: string | null;
  customer_state: string | null;
  customer_zip: string | null;
  config: BuildingConfig;
  pricing: PriceBreakdown;
  subtotal: number;
  tax_rate: number;
  tax_amount: number;
  total: number;
  notes: string | null;
  valid_until: string | null;
  created_at: string;
  updated_at: string;
}
