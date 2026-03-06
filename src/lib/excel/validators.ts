import type { PricingMatrices, StandardMatrices, WidespanMatrices } from "@/types/pricing";

export interface ValidationResult {
  valid: boolean;
  errors: string[];
  warnings: string[];
}

function countKeys(obj: Record<string, unknown>): number {
  return Object.keys(obj).length;
}

function validateStandard(m: StandardMatrices): ValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];

  if (countKeys(m.basePrice) === 0) errors.push("Base price matrix is empty");
  if (countKeys(m.roofStyle) === 0) errors.push("Roof style matrix is empty");
  if (countKeys(m.legs.small) === 0) errors.push("Legs (small) matrix is empty");
  if (countKeys(m.legs.large) === 0) warnings.push("Legs (large) matrix is empty");
  if (countKeys(m.sides) === 0) errors.push("Sides matrix is empty");
  if (countKeys(m.ends) === 0) errors.push("Ends matrix is empty");
  if (countKeys(m.laborEquipment) === 0) errors.push("Labor/Equipment matrix is empty");
  if (countKeys(m.plans) === 0) errors.push("Plans matrix is empty");

  if (countKeys(m.accessories.walkInDoors) === 0) warnings.push("Walk-in doors list is empty");
  if (countKeys(m.accessories.windows) === 0) warnings.push("Windows list is empty");
  if (countKeys(m.changers.widthBuckets) === 0) warnings.push("Width buckets are empty");
  if (countKeys(m.changers.lengthBuckets) === 0) warnings.push("Length buckets are empty");
  if (countKeys(m.snow.trussSpacing) === 0) warnings.push("Truss spacing matrix is empty");
  if (countKeys(m.snow.windThresholdByState) === 0) warnings.push("Wind threshold by state is empty");

  return { valid: errors.length === 0, errors, warnings };
}

function validateWidespan(m: WidespanMatrices): ValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];

  if (countKeys(m.basePrice) === 0) errors.push("Base price matrix is empty");
  if (countKeys(m.legs) === 0) errors.push("Legs matrix is empty");
  if (countKeys(m.sides) === 0) errors.push("Sides matrix is empty");
  if (countKeys(m.ends) === 0) errors.push("Ends matrix is empty");
  if (countKeys(m.laborEquipment) === 0) errors.push("Equipment matrix is empty");
  if (countKeys(m.plans) === 0) errors.push("Plans matrix is empty");

  if (countKeys(m.accessories.walkInDoors) === 0) warnings.push("Walk-in doors list is empty");
  if (countKeys(m.accessories.windows) === 0) warnings.push("Windows list is empty");
  if (countKeys(m.accessories.rollUpDoors) === 0) warnings.push("Roll-up doors list is empty");
  if (countKeys(m.wainscot.sides) === 0) warnings.push("Wainscot sides lookup is empty");
  if (countKeys(m.wainscot.ends) === 0) warnings.push("Wainscot ends lookup is empty");
  if (countKeys(m.changers.widthBuckets) === 0) warnings.push("Width buckets are empty");
  if (countKeys(m.changers.lengthBuckets) === 0) warnings.push("Length buckets are empty");

  return { valid: errors.length === 0, errors, warnings };
}

export function validateMatrices(matrices: PricingMatrices): ValidationResult {
  return matrices.type === "standard"
    ? validateStandard(matrices)
    : validateWidespan(matrices);
}
