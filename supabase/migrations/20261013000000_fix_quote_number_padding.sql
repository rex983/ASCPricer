-- next_quote_number() used LPAD(n, 4, '0'), and LPAD truncates longer input:
-- quote #10000 would render as 'ASC-YYYY-1000', colliding with the existing
-- 1000th quote (asc_quotes.quote_number is UNIQUE) and failing every insert
-- from then on. Keep the 4-digit zero padding, but never truncate.
--
-- Output for 1..9999 is unchanged. Idempotent (CREATE OR REPLACE).

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

  RETURN 'ASC-' || current_year || '-' ||
    LPAD(next_num::TEXT, GREATEST(4, length(next_num::TEXT)), '0');
END;
$$;
