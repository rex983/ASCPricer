// Standard building width buckets (12-30, even increments)
export const STANDARD_WIDTHS = [12, 18, 20, 22, 24, 26, 28, 30] as const;

// Widespan building width buckets (32-60, even increments)
export const WIDESPAN_WIDTHS = [
  32, 34, 36, 38, 40, 42, 44, 46, 48, 50, 52, 54, 56, 58, 60,
] as const;

// Standard length buckets (5ft increments)
export const STANDARD_BASE_LENGTHS = [20, 25, 30, 35, 40, 45, 50] as const;
export const STANDARD_EXTENDED_LENGTHS = [
  55, 60, 65, 70, 75, 80, 85, 90, 95, 100,
] as const;

// Widespan lengths go up to 200
export const WIDESPAN_MAX_LENGTH = 200;

// Height ranges
export const STANDARD_MIN_HEIGHT = 6;
export const STANDARD_MAX_HEIGHT = 20;
export const WIDESPAN_MIN_HEIGHT = 8;
export const WIDESPAN_MAX_HEIGHT = 20;

// Insulation rates
export const FIBERGLASS_RATE = 2.25;
export const THERMAL_RATE = 1.65;

// Sheet metal multipliers (widespan)
export const SHEET_METAL_MULTIPLIERS = {
  "29g_agg": 1.0,
  "26g_agg": 1.1,
  "26g_pbr": 1.2,
} as const;

// Snow engineering height multipliers
export const HEIGHT_MULTIPLIERS = {
  "6-12": 1.0,
  "13-15": 2.0,
  "16-18": 2.5,
  "19-20": 3.0,
} as const;

// Length decomposition for 55-100ft (standard)
// Each extended length = sum of two base lengths
export const LENGTH_DECOMPOSITION: Record<number, [number, number]> = {
  55: [25, 30],
  60: [30, 30],
  65: [30, 35],
  70: [35, 35],
  75: [35, 40],
  80: [40, 40],
  85: [40, 45],
  90: [45, 45],
  95: [45, 50],
  100: [50, 50],
};

// Roof style symbols
export const ROOF_STYLE_KEYS = {
  standard: "STD",
  a_frame_horizontal: "AFH",
  a_frame_vertical: "AFV",
} as const;

// Building type classifications (standard)
export const BUILDING_TYPES = {
  S: "Small",
  M: "Medium",
  T: "Tall",
  ET: "Extra Tall",
} as const;

// Building type classifications (widespan)
export const WIDESPAN_BUILDING_TYPES = {
  S: "Standard",
  G: "Giant",
} as const;

// Wind load categories
export const WIND_LOAD_CATEGORIES = [
  105, 115, 130, 140, 155, 165, 180,
] as const;

// Widespan diagonal bracing
export const WIDESPAN_BRACE_COUNT_SHORT = 4; // length ≤ 50
export const WIDESPAN_BRACE_COUNT_LONG = 6; // length > 50
export const WIDESPAN_BRACE_ENDS_EXTRA = 2; // per enclosed end
export const WIDESPAN_BRACE_PRICE = 350;

// Standard diagonal bracing
export const STANDARD_BRACE_BASE_PRICE = 90;
export const STANDARD_BRACE_TALL_SURCHARGE = 50; // if height > 12
