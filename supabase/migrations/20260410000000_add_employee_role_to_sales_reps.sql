-- Add employee_role to distinguish sales reps, sales managers, and BST (customer service)
ALTER TABLE asc_sales_reps
  ADD COLUMN IF NOT EXISTS employee_role TEXT NOT NULL DEFAULT 'sales_rep'
  CHECK (employee_role IN ('sales_rep', 'sales_manager', 'bst'));

-- Index for filtering by role
CREATE INDEX IF NOT EXISTS idx_asc_sales_reps_employee_role ON asc_sales_reps(employee_role);
