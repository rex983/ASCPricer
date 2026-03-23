import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export function formatCurrency(amount: number): string {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 2,
  }).format(amount);
}

export function formatDate(date: string | Date): string {
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
  }).format(new Date(date));
}

export const STATUS_COLORS: Record<string, "default" | "secondary" | "destructive"> = {
  draft: "secondary",
  sent: "default",
  accepted: "default",
  expired: "destructive",
};

export function canDeleteRecord(role?: string): boolean {
  return role === "admin" || role === "manager" || role === "sales_rep";
}
