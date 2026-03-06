/**
 * Determine building type classification based on width and height.
 *
 * Standard buildings (width ≤ 30):
 *   S (Small): height 6-12, all widths
 *   M (Medium): height 6-12, width 26-30 (reinforced)
 *   T (Tall): height 13-15
 *   ET (Extra Tall): height 16-20
 *
 * Widespan buildings (width 32+):
 *   S (Standard): typical configs
 *   G (Giant): width 48+ or very tall
 */
export function getStandardBuildingType(
  width: number,
  height: number
): "S" | "M" | "T" | "ET" {
  if (height >= 19) return "ET";
  if (height >= 16) return "ET";
  if (height >= 13) return "T";
  if (width >= 26) return "M";
  return "S";
}

export function getWidespanBuildingType(
  width: number,
  height: number
): "S" | "G" {
  if (width >= 48 || height >= 16) return "G";
  return "S";
}

/**
 * Determine if a width falls into standard or widespan category.
 */
export function getSpreadsheetType(width: number): "standard" | "widespan" {
  return width <= 30 ? "standard" : "widespan";
}
