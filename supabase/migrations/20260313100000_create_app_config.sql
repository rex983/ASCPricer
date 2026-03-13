CREATE TABLE IF NOT EXISTS asc_app_config (
  key TEXT PRIMARY KEY,
  value JSONB NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_by TEXT
);

ALTER TABLE asc_app_config ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Anyone can read config" ON asc_app_config FOR SELECT USING (true);

-- Seed with defaults
INSERT INTO asc_app_config (key, value) VALUES
  ('standard_widths', '[12, 18, 20, 22, 24, 26, 28, 30]'),
  ('widespan_widths', '[32, 34, 36, 38, 40, 42, 44, 46, 48, 50, 52, 54, 56, 58, 60]'),
  ('height_range_standard', '{"min": 6, "max": 20}'),
  ('height_range_widespan', '{"min": 8, "max": 20}'),
  ('roof_styles', '[{"value": "standard", "label": "Standard (Regular)"}, {"value": "a_frame_horizontal", "label": "A-Frame Horizontal"}, {"value": "a_frame_vertical", "label": "A-Frame Vertical"}]'),
  ('side_coverage_options', '[{"value": "open", "label": "Open"}, {"value": "3", "label": "3'' Sides Down"}, {"value": "4", "label": "4'' Sides Down"}, {"value": "5", "label": "5'' Sides Down"}, {"value": "6", "label": "6'' Sides Down"}, {"value": "7", "label": "7'' Sides Down"}, {"value": "8", "label": "8'' Sides Down"}, {"value": "9", "label": "9'' Sides Down"}, {"value": "10", "label": "10'' Sides Down"}, {"value": "fully_enclosed", "label": "Fully Enclosed"}]'),
  ('end_types', '[{"value": "enclosed", "label": "Enclosed Ends"}, {"value": "gable", "label": "Gable Ends"}, {"value": "extended_gable", "label": "Extended Gable"}]'),
  ('gauge_options', '[12, 14]'),
  ('orientation_options', '[{"value": "horizontal", "label": "Horizontal"}, {"value": "vertical", "label": "Vertical"}]'),
  ('insulation_types', '[{"value": "none", "label": "None"}, {"value": "fiberglass", "label": "2\" Fiberglass"}, {"value": "thermal", "label": "Thermal"}]'),
  ('insulation_scopes', '[{"value": "none", "label": "None"}, {"value": "roof_only", "label": "Roof Only"}, {"value": "fully_insulated", "label": "Fully Insulated"}]'),
  ('wainscot_options', '[{"value": "none", "label": "None"}, {"value": "full", "label": "Full"}, {"value": "sides", "label": "Sides Only"}, {"value": "ends", "label": "Ends Only"}]'),
  ('plans_snow_surcharges', '{"20LL": 0, "30GL": 0, "27LL": 225, "40GL": 225, "34LL": 300, "50GL": 300, "41LL": 375, "60GL": 375, "47LL": 450, "70GL": 450, "54LL": 525, "80GL": 525, "90GL": 750, "61LL": 900}'),
  ('brace_prices', '{"standard_base": 90, "standard_tall_surcharge": 50, "widespan_short": 4, "widespan_long": 6, "widespan_ends_extra": 2, "widespan_price": 350}'),
  ('disclaimers', '["* IF THERE IS A PRICE DISCREPANCY OVER $20, AMERICAN STEEL CARPORTS INC. RESERVES THE RIGHT TO CANCEL THE ORDER.", "** Plans & Calculations Cost May Vary and are not Final.", "*** Agg Units need to be priced separately.", "Due to Snow Concerns in Northern Areas, it is Highly Recommended to go A-Frame Vertical for Roof Style.", "FOR ANY SNOW / WIND LOADS HIGHER THAN THE LISTED OPTIONS, PLEASE CONTACT THE ENGINEERING DEPARTMENT.", "QUOTE EXCLUDES ANY AND ALL ITEMS NOT SPECIFIED."]')
ON CONFLICT (key) DO NOTHING;
