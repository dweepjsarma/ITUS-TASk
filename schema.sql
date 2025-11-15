-- Create index to speed up company + date lookups
CREATE INDEX IF NOT EXISTS idx_company_date
ON ebitda_margins (accord_code, date);
