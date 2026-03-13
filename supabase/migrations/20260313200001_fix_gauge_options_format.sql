-- Fix gauge_options from number array to value/label format
UPDATE asc_app_config
SET value = '[{"value": "14", "label": "14 Gauge"}, {"value": "12", "label": "12 Gauge"}]'::jsonb,
    updated_at = now(),
    updated_by = 'migration'
WHERE key = 'gauge_options';
