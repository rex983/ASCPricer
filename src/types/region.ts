import type { SpreadsheetType } from "./pricing";

export interface Region {
  id: string;
  name: string;
  slug: string;
  states: string[];
  spreadsheet_type: SpreadsheetType;
  is_active: boolean;
  created_at: string;
  updated_at: string;
}
