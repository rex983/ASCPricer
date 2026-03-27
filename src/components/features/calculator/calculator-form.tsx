"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import { useRouter } from "next/navigation";
import type { BuildingConfig, PricingMatrices, SpreadsheetType } from "@/types/pricing";
import { usePricingEngine } from "@/hooks/use-pricing-engine";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Switch } from "@/components/ui/switch";
import { Textarea } from "@/components/ui/textarea";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { useSession } from "next-auth/react";
import { PriceSummary } from "./price-summary";
import {
  STANDARD_WIDTHS,
  WIDESPAN_WIDTHS,
  STANDARD_MIN_HEIGHT,
  STANDARD_MAX_HEIGHT,
  WIDESPAN_MIN_HEIGHT,
  WIDESPAN_MAX_HEIGHT,
  STANDARD_SNOW_LOAD_OPTIONS,
  WIDESPAN_SNOW_LOAD_OPTIONS,
  DEFAULT_DISCLAIMERS,
} from "@/lib/pricing/constants";
import type { AppConfig } from "@/lib/pricing/constants";

interface CalculatorFormProps {
  spreadsheetType: SpreadsheetType;
  matrices: PricingMatrices | null;
  regionId: string;
  regionStates?: string[];
  appConfig?: AppConfig;
}

type VLItem = { value: string; label: string };

// ── Fallback defaults (used when config not loaded) ──

const DEFAULT_ROOF_OPTIONS: VLItem[] = [
  { value: "standard", label: "Standard (Regular)" },
  { value: "a_frame_horizontal", label: "A-Frame Horizontal" },
  { value: "a_frame_vertical", label: "A-Frame Vertical" },
];

const DEFAULT_SIDE_COVERAGE: VLItem[] = [
  { value: "open", label: "Open" },
  { value: "3", label: "3' Sides Down" },
  { value: "4", label: "4' Sides Down" },
  { value: "5", label: "5' Sides Down" },
  { value: "6", label: "6' Sides Down" },
  { value: "7", label: "7' Sides Down" },
  { value: "8", label: "8' Sides Down" },
  { value: "9", label: "9' Sides Down" },
  { value: "10", label: "10' Sides Down" },
  { value: "fully_enclosed", label: "Fully Enclosed" },
];

const DEFAULT_END_TYPES: VLItem[] = [
  { value: "enclosed", label: "Fully Enclosed" },
  { value: "gable", label: "Gable" },
  { value: "extended_gable", label: "Extended Gable" },
];

const DEFAULT_INSULATION_TYPES: VLItem[] = [
  { value: "none", label: "None" },
  { value: "fiberglass", label: "Fiberglass" },
  { value: "thermal", label: "Thermal" },
];

const DEFAULT_INSULATION_SCOPES: VLItem[] = [
  { value: "none", label: "None" },
  { value: "roof_only", label: "Roof Only" },
  { value: "fully_insulated", label: "Fully Insulated" },
];

const DEFAULT_WAINSCOT: VLItem[] = [
  { value: "none", label: "None" },
  { value: "full", label: "Full (Sides + Ends)" },
  { value: "sides", label: "Sides Only" },
  { value: "ends", label: "Ends Only" },
];

const DEFAULT_SHEET_METAL: VLItem[] = [
  { value: "29g_agg", label: "29G Agg Panel" },
  { value: "26g_agg", label: "26G Agg Panel" },
  { value: "26g_pbr", label: "26G PBR Panel" },
];

const DEFAULT_GAUGE: VLItem[] = [
  { value: "14", label: "14 Gauge" },
  { value: "12", label: "12 Gauge" },
];

const DEFAULT_ORIENTATION: VLItem[] = [
  { value: "horizontal", label: "Horizontal" },
  { value: "vertical", label: "Vertical" },
];

// ── Config helpers ──

function getVLArray(cfg: AppConfig | undefined, key: string, fallback: VLItem[]): VLItem[] {
  const val = cfg?.[key];
  if (Array.isArray(val) && val.length > 0 && typeof val[0] === "object" && "value" in val[0]) {
    return val as VLItem[];
  }
  return fallback;
}

function getNumberArray(cfg: AppConfig | undefined, key: string, fallback: readonly number[]): number[] {
  const val = cfg?.[key];
  if (Array.isArray(val) && val.length > 0 && typeof val[0] === "number") {
    return val as number[];
  }
  return [...fallback];
}

function getHeightRange(cfg: AppConfig | undefined, key: string, defMin: number, defMax: number): { min: number; max: number } {
  const val = cfg?.[key];
  if (val && typeof val === "object" && "min" in (val as Record<string, unknown>)) {
    const r = val as { min: number; max: number };
    return { min: r.min ?? defMin, max: r.max ?? defMax };
  }
  return { min: defMin, max: defMax };
}

function getStringArray(cfg: AppConfig | undefined, key: string, fallback: string[]): string[] {
  const val = cfg?.[key];
  if (Array.isArray(val) && val.length > 0 && typeof val[0] === "string") {
    return val as string[];
  }
  return fallback;
}

function getDefaultConfig(type: SpreadsheetType): BuildingConfig {
  const isWidespan = type === "widespan";
  return {
    width: isWidespan ? 40 : 12,
    length: isWidespan ? 30 : 20,
    height: isWidespan ? 10 : 8,
    gauge: isWidespan ? 12 : 14,
    roofStyle: "a_frame_vertical",
    sheetMetal: isWidespan ? "29g_agg" : undefined,
    sidesCoverage: isWidespan ? "open" : "fully_enclosed",
    sidesOrientation: "horizontal",
    sidesQty: isWidespan ? 0 : 2,
    endType: "enclosed",
    endsOrientation: "horizontal",
    endsQty: isWidespan ? 0 : 2,
    walkInDoorType: undefined,
    walkInDoorQty: 0,
    windowType: undefined,
    windowQty: 0,
    rollUpEndSize1: undefined,
    rollUpEndQty1: 0,
    rollUpEndSize2: undefined,
    rollUpEndQty2: 0,
    rollUpSideSize1: undefined,
    rollUpSideQty1: 0,
    rollUpSideSize2: undefined,
    rollUpSideQty2: 0,
    insulationType: "none",
    insulationScope: "none",
    wainscot: "none",
    snowLoad: isWidespan ? undefined : "30GL",
    windRating: isWidespan ? 90 : 105,
    includePlans: false,
    taxRate: 0,
  };
}

/** Extract door/window option labels from the matrices accessories. */
function getAccessoryOptions(matrices: PricingMatrices | null) {
  if (!matrices) return { doors: [], windows: [], rollUpEnds: [], rollUpSides: [] };

  const acc = matrices.accessories;
  const doors = Object.keys(acc.walkInDoors || {});
  const windows = Object.keys(acc.windows || {});

  // Standard has separate ends/sides roll-up price columns
  // Widespan uses a single rollUpDoors lookup
  // Also handle old schema where standard had rollUpDoors instead of rollUpEnds/rollUpSides
  let rollUpEnds: string[];
  let rollUpSides: string[];
  if (matrices.type === "standard") {
    const stdAcc = matrices.accessories as Record<string, unknown>;
    rollUpEnds = Object.keys(
      (stdAcc.rollUpEnds as Record<string, number>) ||
      (stdAcc.rollUpDoors as Record<string, number>) || {}
    );
    rollUpSides = Object.keys(
      (stdAcc.rollUpSides as Record<string, number>) ||
      (stdAcc.rollUpDoors as Record<string, number>) || {}
    );
  } else {
    rollUpEnds = Object.keys(matrices.accessories.rollUpDoors || {});
    rollUpSides = Object.keys(matrices.accessories.rollUpDoors || {});
  }

  return { doors, windows, rollUpEnds, rollUpSides };
}

export function CalculatorForm({ spreadsheetType, matrices, regionId, regionStates = [], appConfig }: CalculatorFormProps) {
  const isWidespan = spreadsheetType === "widespan";
  const router = useRouter();
  const [config, setConfig] = useState<BuildingConfig>(() => getDefaultConfig(spreadsheetType));
  const [saving, setSaving] = useState(false);
  const breakdown = usePricingEngine(config, matrices);

  // Config-driven options (fall back to hardcoded defaults), memoized on appConfig
  const {
    widths, minHeight, maxHeight, roofOptions, sideCoverageOptions,
    endTypeOptions, insulationTypes, insulationScopes, wainscotOptions,
    gaugeOptions, orientationOptions, snowLoadOptions, sheetMetalOptions,
    disclaimers,
  } = useMemo(() => {
    const hr = getHeightRange(appConfig, isWidespan ? "height_range_widespan" : "height_range_standard", isWidespan ? WIDESPAN_MIN_HEIGHT : STANDARD_MIN_HEIGHT, isWidespan ? WIDESPAN_MAX_HEIGHT : STANDARD_MAX_HEIGHT);
    return {
      widths: getNumberArray(appConfig, isWidespan ? "widespan_widths" : "standard_widths", isWidespan ? WIDESPAN_WIDTHS : STANDARD_WIDTHS),
      minHeight: hr.min,
      maxHeight: hr.max,
      gaugeOptions: getVLArray(appConfig, "gauge_options", DEFAULT_GAUGE),
      roofOptions: getVLArray(appConfig, "roof_styles", DEFAULT_ROOF_OPTIONS),
      sideCoverageOptions: getVLArray(appConfig, "side_coverage_options", DEFAULT_SIDE_COVERAGE),
      endTypeOptions: getVLArray(appConfig, "end_types", DEFAULT_END_TYPES),
      orientationOptions: getVLArray(appConfig, "orientation_options", DEFAULT_ORIENTATION),
      insulationTypes: getVLArray(appConfig, "insulation_types", DEFAULT_INSULATION_TYPES),
      insulationScopes: getVLArray(appConfig, "insulation_scopes", DEFAULT_INSULATION_SCOPES),
      wainscotOptions: getVLArray(appConfig, "wainscot_options", DEFAULT_WAINSCOT),
      snowLoadOptions: getVLArray(appConfig, isWidespan ? "widespan_snow_load_options" : "standard_snow_load_options", isWidespan ? [...WIDESPAN_SNOW_LOAD_OPTIONS] : [...STANDARD_SNOW_LOAD_OPTIONS]),
      sheetMetalOptions: getVLArray(appConfig, "sheet_metal_options", DEFAULT_SHEET_METAL),
      disclaimers: getStringArray(appConfig, "disclaimers", [...DEFAULT_DISCLAIMERS]),
    };
  }, [appConfig, isWidespan]);
  const { doors, windows, rollUpEnds, rollUpSides } = useMemo(() => getAccessoryOptions(matrices), [matrices]);

  // Reset config when spreadsheet type changes
  useEffect(() => {
    setConfig(getDefaultConfig(spreadsheetType));
  }, [spreadsheetType]);

  const update = useCallback(<K extends keyof BuildingConfig>(key: K, value: BuildingConfig[K]) => {
    setConfig((prev) => ({ ...prev, [key]: value }));
  }, []);

  // Sync sidesQty based on coverage
  const handleSidesCoverage = useCallback((value: string) => {
    const isOpen = value === "open";
    setConfig((prev) => ({
      ...prev,
      sidesCoverage: value === "open" ? "open" : value === "fully_enclosed" ? "fully_enclosed" : value,
      sidesQty: isOpen ? 0 : prev.sidesQty === 0 ? 2 : prev.sidesQty,
    }));
  }, []);

  // Sync endsQty when type changes
  const handleEndType = useCallback((value: string) => {
    setConfig((prev) => ({
      ...prev,
      endType: value as BuildingConfig["endType"],
      endsQty: prev.endsQty === 0 ? 2 : prev.endsQty,
    }));
  }, []);

  const [saveDialogOpen, setSaveDialogOpen] = useState(false);
  const { data: session } = useSession();
  const [customerMode, setCustomerMode] = useState<"new" | "existing">("new");
  const [customerForm, setCustomerForm] = useState({
    name: "", email: "", phone: "", address: "", city: "", state: "", zip: "",
  });
  const setCust = (field: string, value: string) =>
    setCustomerForm((prev) => ({ ...prev, [field]: value }));
  const [quoteNotes, setQuoteNotes] = useState("");

  // Existing customer search
  const [customerSearch, setCustomerSearch] = useState("");
  const [customerResults, setCustomerResults] = useState<
    { id: string; name: string; email: string | null; phone: string | null; city: string | null; state: string | null; office: string }[]
  >([]);
  const [selectedCustomer, setSelectedCustomer] = useState<typeof customerResults[number] | null>(null);
  const [searchingCustomers, setSearchingCustomers] = useState(false);

  // Debounced customer search
  useEffect(() => {
    if (customerMode !== "existing" || customerSearch.length < 2) {
      setCustomerResults([]);
      return;
    }
    const timer = setTimeout(async () => {
      setSearchingCustomers(true);
      try {
        const res = await fetch(`/api/customers?search=${encodeURIComponent(customerSearch)}`);
        const data = await res.json();
        if (Array.isArray(data)) setCustomerResults(data);
      } catch { /* ignore */ }
      setSearchingCustomers(false);
    }, 300);
    return () => clearTimeout(timer);
  }, [customerSearch, customerMode]);

  const handleSaveQuote = useCallback(async () => {
    if (!breakdown) return;
    setSaving(true);
    try {
      const payload: Record<string, unknown> = {
        regionId,
        config,
        notes: quoteNotes || undefined,
        office: session?.user?.office || undefined,
      };

      if (customerMode === "existing" && selectedCustomer) {
        payload.customerId = selectedCustomer.id;
      } else if (customerMode === "new") {
        payload.customer = {
          name: customerForm.name || undefined,
          email: customerForm.email || undefined,
          phone: customerForm.phone || undefined,
          address: customerForm.address || undefined,
          city: customerForm.city || undefined,
          state: customerForm.state || undefined,
          zip: customerForm.zip || undefined,
        };
      }

      const res = await fetch("/api/quotes", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      if (data.id) {
        router.push(`/quotes/${data.id}`);
      } else {
        alert(data.error || "Failed to save quote");
      }
    } catch {
      alert("Network error saving quote");
    } finally {
      setSaving(false);
    }
  }, [breakdown, regionId, config, customerForm, customerMode, selectedCustomer, quoteNotes, session, router]);

  return (
    <div className="grid gap-4 lg:grid-cols-[1fr_320px]">
      {/* Left: Form */}
      <div className="space-y-4">
        {/* ── Dimensions + Roof/Sheet Metal ── */}
        <Card>
          <CardContent className="pt-4 pb-3 space-y-3">
            <p className="text-xs font-medium text-muted-foreground uppercase tracking-wide">Dimensions</p>
            <div className="grid grid-cols-2 gap-x-4 gap-y-3 sm:grid-cols-3 lg:grid-cols-5">
              <div className="space-y-1">
                <Label className="text-xs">Width</Label>
                <Select value={String(config.width)} onValueChange={(v) => update("width", Number(v))}>
                  <SelectTrigger className="h-8 text-sm"><SelectValue /></SelectTrigger>
                  <SelectContent>
                    {widths.map((w) => (
                      <SelectItem key={w} value={String(w)}>{w}&apos;</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <div className="space-y-1">
                <Label className="text-xs">Length</Label>
                <Input className="h-8 text-sm" type="number" min={20} max={isWidespan ? 200 : 100} step={1}
                  value={config.length} onChange={(e) => update("length", Number(e.target.value))} />
              </div>
              <div className="space-y-1">
                <Label className="text-xs">Height</Label>
                <Input className="h-8 text-sm" type="number" min={minHeight} max={maxHeight} step={1}
                  value={config.height} onChange={(e) => update("height", Number(e.target.value))} />
              </div>
              <div className="space-y-1">
                <Label className="text-xs">Gauge</Label>
                <Select value={String(config.gauge)} onValueChange={(v) => update("gauge", Number(v) as 12 | 14)}>
                  <SelectTrigger className="h-8 text-sm"><SelectValue /></SelectTrigger>
                  <SelectContent>
                    {gaugeOptions.map((opt) => (
                      <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <div className="space-y-1">
                <Label className="text-xs">{isWidespan ? "Sheet Metal" : "Roof Style"}</Label>
                {isWidespan ? (
                  <Select value={config.sheetMetal || "29g_agg"} onValueChange={(v) => update("sheetMetal", v as BuildingConfig["sheetMetal"])}>
                    <SelectTrigger className="h-8 text-sm"><SelectValue /></SelectTrigger>
                    <SelectContent>
                      {sheetMetalOptions.map((opt) => (
                        <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                ) : (
                  <Select value={config.roofStyle} onValueChange={(v) => update("roofStyle", v as BuildingConfig["roofStyle"])}>
                    <SelectTrigger className="h-8 text-sm"><SelectValue /></SelectTrigger>
                    <SelectContent>
                      {roofOptions.map((opt) => (
                        <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                )}
              </div>
            </div>
          </CardContent>
        </Card>

        {/* ── Sides & Ends ── */}
        <Card>
          <CardContent className="pt-4 pb-3 space-y-3">
            <p className="text-xs font-medium text-muted-foreground uppercase tracking-wide">Sides & Ends</p>
            <div className="grid grid-cols-2 gap-x-4 gap-y-3 sm:grid-cols-3">
              <div className="space-y-1">
                <Label className="text-xs">Side Coverage</Label>
                <Select
                  value={config.sidesCoverage === "open" ? "open" : config.sidesCoverage === "fully_enclosed" ? "fully_enclosed" : config.sidesCoverage}
                  onValueChange={handleSidesCoverage}
                >
                  <SelectTrigger className="h-8 text-sm"><SelectValue /></SelectTrigger>
                  <SelectContent>
                    {sideCoverageOptions.map((opt) => (
                      <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              {config.sidesCoverage !== "open" && (
                <div className="space-y-1">
                  <Label className="text-xs"># Sides</Label>
                  <Select value={String(config.sidesQty)} onValueChange={(v) => update("sidesQty", Number(v) as 0 | 1 | 2)}>
                    <SelectTrigger className="h-8 text-sm"><SelectValue /></SelectTrigger>
                    <SelectContent>
                      <SelectItem value="1">1</SelectItem>
                      <SelectItem value="2">2</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
              )}
              {config.sidesCoverage !== "open" && !isWidespan && (
                <div className="space-y-1">
                  <Label className="text-xs">Side Panel</Label>
                  <Select value={config.sidesOrientation} onValueChange={(v) => update("sidesOrientation", v as "horizontal" | "vertical")}>
                    <SelectTrigger className="h-8 text-sm"><SelectValue /></SelectTrigger>
                    <SelectContent>
                      {orientationOptions.map((opt) => (
                        <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              )}
              {config.endsQty > 0 ? (
                <div className="space-y-1">
                  <Label className="text-xs">End Type</Label>
                  <Select value={config.endType} onValueChange={handleEndType}>
                    <SelectTrigger className="h-8 text-sm"><SelectValue /></SelectTrigger>
                    <SelectContent>
                      {endTypeOptions.filter((opt) => isWidespan ? opt.value !== "extended_gable" : true).map((opt) => (
                        <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              ) : (
                <div />
              )}
              <div className="space-y-1">
                <Label className="text-xs"># Ends</Label>
                <Select value={String(config.endsQty)} onValueChange={(v) => update("endsQty", Number(v) as 0 | 1 | 2)}>
                  <SelectTrigger className="h-8 text-sm"><SelectValue /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="0">None</SelectItem>
                    <SelectItem value="1">1</SelectItem>
                    <SelectItem value="2">2</SelectItem>
                  </SelectContent>
                </Select>
              </div>
              {config.endsQty > 0 && !isWidespan && (
                <div className="space-y-1">
                  <Label className="text-xs">End Panel</Label>
                  <Select value={config.endsOrientation} onValueChange={(v) => update("endsOrientation", v as "horizontal" | "vertical")}>
                    <SelectTrigger className="h-8 text-sm"><SelectValue /></SelectTrigger>
                    <SelectContent>
                      {orientationOptions.map((opt) => (
                        <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              )}
            </div>
          </CardContent>
        </Card>

        {/* ── Doors & Windows ── */}
        <Card>
          <CardContent className="pt-4 pb-3 space-y-3">
            <p className="text-xs font-medium text-muted-foreground uppercase tracking-wide">Doors & Windows</p>
            <div className="grid grid-cols-2 gap-x-4 gap-y-3 sm:grid-cols-3">
              <div className="space-y-1">
                <Label className="text-xs">Walk-In Door</Label>
                <Select
                  value={config.walkInDoorType || "__none__"}
                  onValueChange={(v) => {
                    const val = v === "__none__" ? undefined : v;
                    setConfig((prev) => ({
                      ...prev,
                      walkInDoorType: val,
                      walkInDoorQty: val ? Math.max(1, prev.walkInDoorQty) : 0,
                    }));
                  }}
                >
                  <SelectTrigger className="h-8 text-sm"><SelectValue placeholder="None" /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="__none__">None</SelectItem>
                    {doors.map((d) => (
                      <SelectItem key={d} value={d}>{d}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              {config.walkInDoorType && (
                <div className="space-y-1">
                  <Label className="text-xs">Door Qty</Label>
                  <Input className="h-8 text-sm" type="number" min={1} max={10}
                    value={config.walkInDoorQty} onChange={(e) => update("walkInDoorQty", Number(e.target.value))} />
                </div>
              )}
              <div className="space-y-1">
                <Label className="text-xs">Window</Label>
                <Select
                  value={config.windowType || "__none__"}
                  onValueChange={(v) => {
                    const val = v === "__none__" ? undefined : v;
                    setConfig((prev) => ({
                      ...prev,
                      windowType: val,
                      windowQty: val ? Math.max(1, prev.windowQty) : 0,
                    }));
                  }}
                >
                  <SelectTrigger className="h-8 text-sm"><SelectValue placeholder="None" /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="__none__">None</SelectItem>
                    {windows.map((w) => (
                      <SelectItem key={w} value={w}>{w}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              {config.windowType && (
                <div className="space-y-1">
                  <Label className="text-xs">Win Qty</Label>
                  <Input className="h-8 text-sm" type="number" min={1} max={20}
                    value={config.windowQty} onChange={(e) => update("windowQty", Number(e.target.value))} />
                </div>
              )}
            </div>

            {/* Roll-Up Doors */}
            <div className="grid grid-cols-2 gap-3">
              <div className="space-y-1.5">
                <Label className="text-xs">Roll-Up Doors (Ends)</Label>
              <div className="grid grid-cols-[1fr_60px] gap-1.5">
                <Select
                  value={config.rollUpEndSize1 || "__none__"}
                  onValueChange={(v) => {
                    const val = v === "__none__" ? undefined : v;
                    setConfig((prev) => ({
                      ...prev,
                      rollUpEndSize1: val,
                      rollUpEndQty1: val ? Math.max(1, prev.rollUpEndQty1) : 0,
                    }));
                  }}
                >
                  <SelectTrigger className="h-8 text-sm"><SelectValue placeholder="Size" /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="__none__">None</SelectItem>
                    {rollUpEnds.map((r) => (
                      <SelectItem key={r} value={r}>{r}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
                {config.rollUpEndSize1 ? (
                  <Input className="h-8 text-sm" type="number" min={1} max={10} value={config.rollUpEndQty1}
                    onChange={(e) => update("rollUpEndQty1", Number(e.target.value))} />
                ) : <div />}
              </div>
              <div className="grid grid-cols-[1fr_60px] gap-1.5">
                <Select
                  value={config.rollUpEndSize2 || "__none__"}
                  onValueChange={(v) => {
                    const val = v === "__none__" ? undefined : v;
                    setConfig((prev) => ({
                      ...prev,
                      rollUpEndSize2: val,
                      rollUpEndQty2: val ? Math.max(1, prev.rollUpEndQty2) : 0,
                    }));
                  }}
                >
                  <SelectTrigger className="h-8 text-sm"><SelectValue placeholder="Size" /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="__none__">None</SelectItem>
                    {rollUpEnds.map((r) => (
                      <SelectItem key={`e2-${r}`} value={r}>{r}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
                {config.rollUpEndSize2 ? (
                  <Input className="h-8 text-sm" type="number" min={1} max={10} value={config.rollUpEndQty2}
                    onChange={(e) => update("rollUpEndQty2", Number(e.target.value))} />
                ) : <div />}
              </div>
            </div>

              <div className="space-y-1.5">
                <Label className="text-xs">Roll-Up Doors (Sides)</Label>
              <div className="grid grid-cols-[1fr_60px] gap-1.5">
                <Select
                  value={config.rollUpSideSize1 || "__none__"}
                  onValueChange={(v) => {
                    const val = v === "__none__" ? undefined : v;
                    setConfig((prev) => ({
                      ...prev,
                      rollUpSideSize1: val,
                      rollUpSideQty1: val ? Math.max(1, prev.rollUpSideQty1) : 0,
                    }));
                  }}
                >
                  <SelectTrigger className="h-8 text-sm"><SelectValue placeholder="Size" /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="__none__">None</SelectItem>
                    {rollUpSides.map((r) => (
                      <SelectItem key={`s1-${r}`} value={r}>{r}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
                {config.rollUpSideSize1 ? (
                  <Input className="h-8 text-sm" type="number" min={1} max={10} value={config.rollUpSideQty1}
                    onChange={(e) => update("rollUpSideQty1", Number(e.target.value))} />
                ) : <div />}
              </div>
              <div className="grid grid-cols-[1fr_60px] gap-1.5">
                <Select
                  value={config.rollUpSideSize2 || "__none__"}
                  onValueChange={(v) => {
                    const val = v === "__none__" ? undefined : v;
                    setConfig((prev) => ({
                      ...prev,
                      rollUpSideSize2: val,
                      rollUpSideQty2: val ? Math.max(1, prev.rollUpSideQty2) : 0,
                    }));
                  }}
                >
                  <SelectTrigger className="h-8 text-sm"><SelectValue placeholder="Size" /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="__none__">None</SelectItem>
                    {rollUpSides.map((r) => (
                      <SelectItem key={`s2-${r}`} value={r}>{r}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
                {config.rollUpSideSize2 ? (
                  <Input className="h-8 text-sm" type="number" min={1} max={10} value={config.rollUpSideQty2}
                    onChange={(e) => update("rollUpSideQty2", Number(e.target.value))} />
                ) : <div />}
              </div>
              </div>
            </div>
          </CardContent>
        </Card>

        {/* ── Insulation, Engineering & Tax ── */}
        <Card>
          <CardContent className="pt-4 pb-3 space-y-3">
            <p className="text-xs font-medium text-muted-foreground uppercase tracking-wide">Insulation, Engineering & Tax</p>
            <div className="grid grid-cols-2 gap-x-4 gap-y-3 sm:grid-cols-3">
              <div className="space-y-1">
                <Label className="text-xs">Insulation</Label>
                <Select
                  value={config.insulationType}
                  onValueChange={(v) => {
                    const type = v as BuildingConfig["insulationType"];
                    setConfig((prev) => ({
                      ...prev,
                      insulationType: type,
                      insulationScope: type === "none" ? "none" : prev.insulationScope === "none" ? "fully_insulated" : prev.insulationScope,
                    }));
                  }}
                >
                  <SelectTrigger className="h-8 text-sm"><SelectValue /></SelectTrigger>
                  <SelectContent>
                    {insulationTypes.map((opt) => (
                      <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              {config.insulationType !== "none" && (() => {
                const isAFV = config.roofStyle === "a_frame_vertical" || isWidespan;
                const allVertical = isAFV
                  && config.sidesOrientation === "vertical"
                  && config.endsOrientation === "vertical";
                return (
                  <div className="space-y-1">
                    <Label className="text-xs">Scope</Label>
                    <Select
                      value={config.insulationScope}
                      onValueChange={(v) => update("insulationScope", v as BuildingConfig["insulationScope"])}
                    >
                      <SelectTrigger className="h-8 text-sm"><SelectValue /></SelectTrigger>
                      <SelectContent>
                        {insulationScopes.map((opt) => (
                          <SelectItem
                            key={opt.value}
                            value={opt.value}
                            disabled={
                              (opt.value === "roof_only" && !isAFV) ||
                              (opt.value === "fully_insulated" && !allVertical)
                            }
                          >
                            {opt.label}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>
                );
              })()}

              {isWidespan && (
                <div className="space-y-1">
                  <Label className="text-xs">Wainscot</Label>
                  <Select value={config.wainscot || "none"} onValueChange={(v) => update("wainscot", v as BuildingConfig["wainscot"])}>
                    <SelectTrigger className="h-8 text-sm"><SelectValue /></SelectTrigger>
                    <SelectContent>
                      {wainscotOptions.map((opt) => (
                        <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              )}

              <div className="space-y-1">
                <Label className="text-xs">Snow Load</Label>
                <Select value={config.snowLoad || "__none__"} onValueChange={(v) => update("snowLoad", v === "__none__" ? undefined : v)}>
                  <SelectTrigger className="h-8 text-sm"><SelectValue placeholder="None" /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="__none__">None</SelectItem>
                    {snowLoadOptions.map((opt) => (
                      <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-1">
                <Label className="text-xs">Wind (MPH)</Label>
                <Input className="h-8 text-sm" type="number" min={90} max={180} step={5}
                  value={config.windRating} onChange={(e) => update("windRating", Number(e.target.value))} />
              </div>

              <div className="space-y-1">
                <Label className="text-xs">Tax (%)</Label>
                <Input className="h-8 text-sm" type="number" min={0} max={15} step={0.01}
                  value={Number((config.taxRate * 100).toFixed(4))}
                  onChange={(e) => update("taxRate", Number(e.target.value) / 100)} />
              </div>
            </div>

            <div className="grid grid-cols-2 gap-x-4 gap-y-3 sm:grid-cols-3">
              {!isWidespan && regionStates.length > 0 && (
                <div className="space-y-1">
                  <Label className="text-xs">State</Label>
                  <Select value={config.state || "__none__"} onValueChange={(v) => update("state", v === "__none__" ? undefined : v)}>
                    <SelectTrigger className="h-8 text-sm"><SelectValue placeholder="State" /></SelectTrigger>
                    <SelectContent>
                      <SelectItem value="__none__">None</SelectItem>
                      {regionStates.map((st) => (
                        <SelectItem key={st} value={st}>{st}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              )}

              {isWidespan && (
                <div className="space-y-1">
                  <Label className="text-xs">Anchors</Label>
                  <Select value={config.anchorType || "none"} onValueChange={(v) => update("anchorType", v)}>
                    <SelectTrigger className="h-8 text-sm"><SelectValue /></SelectTrigger>
                    <SelectContent>
                      <SelectItem value="none">None</SelectItem>
                      <SelectItem value="titen_hd">Titen HD Concrete Screws</SelectItem>
                      <SelectItem value="concrete">Concrete Anchors</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
              )}

              <div className="flex items-center gap-2 pt-4">
                <Switch
                  checked={config.includePlans ?? false}
                  onCheckedChange={(v) => update("includePlans", v)}
                />
                <Label className="text-xs">Include Plans</Label>
              </div>
            </div>
          </CardContent>
        </Card>
      </div>

      {/* Right: Price Summary (sticky) */}
      <div className="lg:sticky lg:top-4 lg:self-start space-y-3">
        <Card>
          <CardContent className="pt-4">
            <PriceSummary breakdown={breakdown} isWidespan={isWidespan} disclaimers={disclaimers} />
            {breakdown && (
              <>
                <Button
                  className="mt-4 w-full"
                  onClick={() => setSaveDialogOpen(true)}
                >
                  Save Quote
                </Button>
                <Dialog open={saveDialogOpen} onOpenChange={(open) => {
                  setSaveDialogOpen(open);
                  if (!open) {
                    setSelectedCustomer(null);
                    setCustomerSearch("");
                    setCustomerResults([]);
                  }
                }}>
                  <DialogContent className="max-w-md max-h-[85vh] overflow-y-auto">
                    <DialogHeader>
                      <DialogTitle>Save Quote</DialogTitle>
                    </DialogHeader>
                    <div className="space-y-4">
                      {/* Mode Toggle */}
                      <div className="flex gap-2">
                        <Button
                          variant={customerMode === "new" ? "default" : "outline"}
                          size="sm"
                          className="flex-1"
                          onClick={() => { setCustomerMode("new"); setSelectedCustomer(null); }}
                        >
                          New Customer
                        </Button>
                        <Button
                          variant={customerMode === "existing" ? "default" : "outline"}
                          size="sm"
                          className="flex-1"
                          onClick={() => setCustomerMode("existing")}
                        >
                          Existing Customer
                        </Button>
                      </div>

                      {customerMode === "existing" ? (
                        <div className="space-y-3">
                          {/* Search */}
                          <div className="space-y-1">
                            <Label>Search Customers</Label>
                            <Input
                              value={customerSearch}
                              onChange={(e) => { setCustomerSearch(e.target.value); setSelectedCustomer(null); }}
                              placeholder="Type name, email, or phone..."
                            />
                          </div>

                          {/* Selected customer card */}
                          {selectedCustomer && (
                            <div className="rounded-md border border-primary bg-primary/5 p-3 text-sm space-y-1">
                              <div className="flex items-center justify-between">
                                <span className="font-semibold">{selectedCustomer.name}</span>
                                <span className="text-xs text-muted-foreground font-mono">
                                  {selectedCustomer.id.slice(0, 8)}
                                </span>
                              </div>
                              {selectedCustomer.email && <p className="text-muted-foreground">{selectedCustomer.email}</p>}
                              {selectedCustomer.phone && <p className="text-muted-foreground">{selectedCustomer.phone}</p>}
                              {(selectedCustomer.city || selectedCustomer.state) && (
                                <p className="text-muted-foreground">
                                  {[selectedCustomer.city, selectedCustomer.state].filter(Boolean).join(", ")}
                                </p>
                              )}
                            </div>
                          )}

                          {/* Search results */}
                          {!selectedCustomer && customerSearch.length >= 2 && (
                            <div className="rounded-md border max-h-48 overflow-y-auto">
                              {searchingCustomers ? (
                                <p className="p-3 text-sm text-muted-foreground text-center">Searching...</p>
                              ) : customerResults.length === 0 ? (
                                <p className="p-3 text-sm text-muted-foreground text-center">No customers found</p>
                              ) : (
                                customerResults.map((c) => (
                                  <button
                                    key={c.id}
                                    type="button"
                                    className="w-full text-left px-3 py-2 hover:bg-accent border-b last:border-b-0 text-sm"
                                    onClick={() => setSelectedCustomer(c)}
                                  >
                                    <div className="flex items-center justify-between">
                                      <span className="font-medium">{c.name}</span>
                                      <span className="text-xs text-muted-foreground font-mono">{c.id.slice(0, 8)}</span>
                                    </div>
                                    <p className="text-muted-foreground text-xs">
                                      {[c.email, c.phone, [c.city, c.state].filter(Boolean).join(", ")].filter(Boolean).join(" · ")}
                                    </p>
                                  </button>
                                ))
                              )}
                            </div>
                          )}

                          {/* Notes */}
                          <div className="space-y-1">
                            <Label>Notes</Label>
                            <Textarea
                              value={quoteNotes}
                              onChange={(e) => setQuoteNotes(e.target.value)}
                              placeholder="Any additional notes..."
                            />
                          </div>
                        </div>
                      ) : (
                        /* New Customer Form */
                        <div className="space-y-3">
                          <div className="space-y-1">
                            <Label>Customer Name</Label>
                            <Input
                              value={customerForm.name}
                              onChange={(e) => setCust("name", e.target.value)}
                              placeholder="John Smith"
                            />
                          </div>
                          <div className="grid grid-cols-2 gap-3">
                            <div className="space-y-1">
                              <Label>Email</Label>
                              <Input
                                type="email"
                                value={customerForm.email}
                                onChange={(e) => setCust("email", e.target.value)}
                              />
                            </div>
                            <div className="space-y-1">
                              <Label>Phone</Label>
                              <Input
                                value={customerForm.phone}
                                onChange={(e) => setCust("phone", e.target.value)}
                              />
                            </div>
                          </div>
                          <div className="space-y-1">
                            <Label>Address</Label>
                            <Input
                              value={customerForm.address}
                              onChange={(e) => setCust("address", e.target.value)}
                            />
                          </div>
                          <div className="grid grid-cols-3 gap-3">
                            <div className="space-y-1">
                              <Label>City</Label>
                              <Input
                                value={customerForm.city}
                                onChange={(e) => setCust("city", e.target.value)}
                              />
                            </div>
                            <div className="space-y-1">
                              <Label>State</Label>
                              <Input
                                value={customerForm.state}
                                onChange={(e) => setCust("state", e.target.value)}
                              />
                            </div>
                            <div className="space-y-1">
                              <Label>ZIP</Label>
                              <Input
                                value={customerForm.zip}
                                onChange={(e) => setCust("zip", e.target.value)}
                              />
                            </div>
                          </div>
                          <div className="space-y-1">
                            <Label>Notes</Label>
                            <Textarea
                              value={quoteNotes}
                              onChange={(e) => setQuoteNotes(e.target.value)}
                              placeholder="Any additional notes..."
                            />
                          </div>
                        </div>
                      )}

                      <Button
                        className="w-full"
                        onClick={() => handleSaveQuote()}
                        disabled={saving || (customerMode === "existing" && !selectedCustomer)}
                      >
                        {saving ? "Saving..." : "Confirm & Save"}
                      </Button>
                    </div>
                  </DialogContent>
                </Dialog>
              </>
            )}
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
