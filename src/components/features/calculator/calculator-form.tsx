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
} from "@/lib/pricing/constants";

interface CalculatorFormProps {
  spreadsheetType: SpreadsheetType;
  matrices: PricingMatrices | null;
  regionId: string;
  regionStates?: string[];
}

const STANDARD_ROOF_OPTIONS = [
  { value: "standard", label: "Standard (Regular)" },
  { value: "a_frame_horizontal", label: "A-Frame Horizontal" },
  { value: "a_frame_vertical", label: "A-Frame Vertical" },
];

const SIDE_COVERAGE_OPTIONS = [
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

const END_TYPE_OPTIONS = [
  { value: "enclosed", label: "Fully Enclosed" },
  { value: "gable", label: "Gable" },
  { value: "extended_gable", label: "Extended Gable" },
];

const INSULATION_OPTIONS = [
  { value: "none", label: "None" },
  { value: "fiberglass", label: "Fiberglass" },
  { value: "thermal", label: "Thermal" },
];

const SHEET_METAL_OPTIONS = [
  { value: "29g_agg", label: "29G Agg Panel" },
  { value: "26g_agg", label: "26G Agg Panel" },
  { value: "26g_pbr", label: "26G PBR Panel" },
];

const WAINSCOT_OPTIONS = [
  { value: "none", label: "None" },
  { value: "full", label: "Full (Sides + Ends)" },
  { value: "sides", label: "Sides Only" },
  { value: "ends", label: "Ends Only" },
];

function getDefaultConfig(type: SpreadsheetType): BuildingConfig {
  const isWidespan = type === "widespan";
  return {
    width: isWidespan ? 40 : 24,
    length: 30,
    height: isWidespan ? 10 : 10,
    gauge: isWidespan ? 12 : 14,
    roofStyle: isWidespan ? "a_frame_vertical" : "a_frame_vertical",
    sheetMetal: isWidespan ? "29g_agg" : undefined,
    sidesCoverage: "open",
    sidesOrientation: "horizontal",
    sidesQty: 0,
    endType: "enclosed",
    endsOrientation: "horizontal",
    endsQty: 0,
    walkInDoorType: undefined,
    walkInDoorQty: 0,
    windowType: undefined,
    windowQty: 0,
    rollUpEndSize: undefined,
    rollUpEndQty: 0,
    rollUpSideSize: undefined,
    rollUpSideQty: 0,
    insulationType: "none",
    insulationQty: 0,
    wainscot: "none",
    windRating: 90,
    permitRequired: false,
    diagonalBracing: false,
    includePlans: false,
    taxRate: 0,
  };
}

/** Extract door/window option labels from the matrices accessories. */
function getAccessoryOptions(matrices: PricingMatrices | null) {
  if (!matrices) return { doors: [], windows: [], rollUps: [] };

  const acc = matrices.type === "standard"
    ? matrices.accessories
    : matrices.accessories;

  const doors = Object.keys(acc.walkInDoors || {});
  const windows = Object.keys(acc.windows || {});
  const rollUps = Object.keys(acc.rollUpDoors || {});

  return { doors, windows, rollUps };
}

export function CalculatorForm({ spreadsheetType, matrices, regionId, regionStates = [] }: CalculatorFormProps) {
  const isWidespan = spreadsheetType === "widespan";
  const router = useRouter();
  const [config, setConfig] = useState<BuildingConfig>(() => getDefaultConfig(spreadsheetType));
  const [saving, setSaving] = useState(false);
  const breakdown = usePricingEngine(config, matrices);
  const widths = isWidespan ? WIDESPAN_WIDTHS : STANDARD_WIDTHS;
  const minHeight = isWidespan ? WIDESPAN_MIN_HEIGHT : STANDARD_MIN_HEIGHT;
  const maxHeight = isWidespan ? WIDESPAN_MAX_HEIGHT : STANDARD_MAX_HEIGHT;
  const { doors, windows, rollUps } = useMemo(() => getAccessoryOptions(matrices), [matrices]);

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

  const handleSaveQuote = useCallback(async () => {
    if (!breakdown) return;
    setSaving(true);
    try {
      const res = await fetch("/api/quotes", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ regionId, config }),
      });
      const data = await res.json();
      if (data.id) {
        router.push(`/quotes/${data.id}`);
      }
    } catch {
      // ignore
    } finally {
      setSaving(false);
    }
  }, [breakdown, regionId, config, router]);

  return (
    <div className="grid gap-6 lg:grid-cols-[1fr_320px]">
      {/* Left: Form */}
      <div className="space-y-6">
        {/* ── Dimensions ── */}
        <Card>
          <CardHeader className="pb-4">
            <CardTitle className="text-base">Dimensions</CardTitle>
          </CardHeader>
          <CardContent className="grid grid-cols-2 gap-4 sm:grid-cols-4">
            <div className="space-y-2">
              <Label>Width</Label>
              <Select
                value={String(config.width)}
                onValueChange={(v) => update("width", Number(v))}
              >
                <SelectTrigger><SelectValue /></SelectTrigger>
                <SelectContent>
                  {widths.map((w) => (
                    <SelectItem key={w} value={String(w)}>{w}&apos;</SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <div className="space-y-2">
              <Label>Length (ft)</Label>
              <Input
                type="number"
                min={20}
                max={isWidespan ? 200 : 100}
                step={1}
                value={config.length}
                onChange={(e) => update("length", Number(e.target.value))}
              />
            </div>

            <div className="space-y-2">
              <Label>Height (ft)</Label>
              <Input
                type="number"
                min={minHeight}
                max={maxHeight}
                step={1}
                value={config.height}
                onChange={(e) => update("height", Number(e.target.value))}
              />
            </div>

            <div className="space-y-2">
              <Label>Gauge</Label>
              <Select
                value={String(config.gauge)}
                onValueChange={(v) => update("gauge", Number(v) as 12 | 14)}
              >
                <SelectTrigger><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="14">14 Gauge</SelectItem>
                  <SelectItem value="12">12 Gauge</SelectItem>
                </SelectContent>
              </Select>
            </div>
          </CardContent>
        </Card>

        {/* ── Roof & Sheet Metal ── */}
        <Card>
          <CardHeader className="pb-4">
            <CardTitle className="text-base">
              {isWidespan ? "Sheet Metal" : "Roof Style"}
            </CardTitle>
          </CardHeader>
          <CardContent className="grid grid-cols-2 gap-4">
            {!isWidespan && (
              <div className="space-y-2">
                <Label>Roof Style</Label>
                <Select
                  value={config.roofStyle}
                  onValueChange={(v) => update("roofStyle", v as BuildingConfig["roofStyle"])}
                >
                  <SelectTrigger><SelectValue /></SelectTrigger>
                  <SelectContent>
                    {STANDARD_ROOF_OPTIONS.map((opt) => (
                      <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            )}

            {isWidespan && (
              <div className="space-y-2">
                <Label>Sheet Metal</Label>
                <Select
                  value={config.sheetMetal || "29g_agg"}
                  onValueChange={(v) => update("sheetMetal", v as BuildingConfig["sheetMetal"])}
                >
                  <SelectTrigger><SelectValue /></SelectTrigger>
                  <SelectContent>
                    {SHEET_METAL_OPTIONS.map((opt) => (
                      <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            )}
          </CardContent>
        </Card>

        {/* ── Sides ── */}
        <Card>
          <CardHeader className="pb-4">
            <CardTitle className="text-base">Sides</CardTitle>
          </CardHeader>
          <CardContent className="grid grid-cols-2 gap-4 sm:grid-cols-3">
            <div className="space-y-2">
              <Label>Coverage</Label>
              <Select
                value={config.sidesCoverage === "open" ? "open" : config.sidesCoverage === "fully_enclosed" ? "fully_enclosed" : config.sidesCoverage}
                onValueChange={handleSidesCoverage}
              >
                <SelectTrigger><SelectValue /></SelectTrigger>
                <SelectContent>
                  {SIDE_COVERAGE_OPTIONS.map((opt) => (
                    <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            {config.sidesCoverage !== "open" && (
              <>
                <div className="space-y-2">
                  <Label>Number of Sides</Label>
                  <Select
                    value={String(config.sidesQty)}
                    onValueChange={(v) => update("sidesQty", Number(v) as 0 | 1 | 2)}
                  >
                    <SelectTrigger><SelectValue /></SelectTrigger>
                    <SelectContent>
                      <SelectItem value="1">1 Side</SelectItem>
                      <SelectItem value="2">2 Sides</SelectItem>
                    </SelectContent>
                  </Select>
                </div>

                {!isWidespan && (
                  <div className="space-y-2">
                    <Label>Panel Orientation</Label>
                    <Select
                      value={config.sidesOrientation}
                      onValueChange={(v) => update("sidesOrientation", v as "horizontal" | "vertical")}
                    >
                      <SelectTrigger><SelectValue /></SelectTrigger>
                      <SelectContent>
                        <SelectItem value="horizontal">Horizontal</SelectItem>
                        <SelectItem value="vertical">Vertical</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>
                )}
              </>
            )}
          </CardContent>
        </Card>

        {/* ── Ends ── */}
        <Card>
          <CardHeader className="pb-4">
            <CardTitle className="text-base">Ends</CardTitle>
          </CardHeader>
          <CardContent className="grid grid-cols-2 gap-4 sm:grid-cols-3">
            <div className="space-y-2">
              <Label>Number of Ends</Label>
              <Select
                value={String(config.endsQty)}
                onValueChange={(v) => {
                  const qty = Number(v) as 0 | 1 | 2;
                  update("endsQty", qty);
                }}
              >
                <SelectTrigger><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="0">None</SelectItem>
                  <SelectItem value="1">1 End</SelectItem>
                  <SelectItem value="2">2 Ends</SelectItem>
                </SelectContent>
              </Select>
            </div>

            {config.endsQty > 0 && (
              <>
                <div className="space-y-2">
                  <Label>End Type</Label>
                  <Select
                    value={config.endType}
                    onValueChange={handleEndType}
                  >
                    <SelectTrigger><SelectValue /></SelectTrigger>
                    <SelectContent>
                      {END_TYPE_OPTIONS.filter(
                        (opt) => isWidespan ? opt.value !== "extended_gable" : true
                      ).map((opt) => (
                        <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                {!isWidespan && (
                  <div className="space-y-2">
                    <Label>Panel Orientation</Label>
                    <Select
                      value={config.endsOrientation}
                      onValueChange={(v) => update("endsOrientation", v as "horizontal" | "vertical")}
                    >
                      <SelectTrigger><SelectValue /></SelectTrigger>
                      <SelectContent>
                        <SelectItem value="horizontal">Horizontal</SelectItem>
                        <SelectItem value="vertical">Vertical</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>
                )}
              </>
            )}
          </CardContent>
        </Card>

        {/* ── Doors & Windows ── */}
        <Card>
          <CardHeader className="pb-4">
            <CardTitle className="text-base">Doors & Windows</CardTitle>
          </CardHeader>
          <CardContent className="space-y-4">
            {/* Walk-In Doors */}
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-2">
                <Label>Walk-In Door</Label>
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
                  <SelectTrigger><SelectValue placeholder="None" /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="__none__">None</SelectItem>
                    {doors.map((d) => (
                      <SelectItem key={d} value={d}>{d}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              {config.walkInDoorType && (
                <div className="space-y-2">
                  <Label>Qty</Label>
                  <Input
                    type="number"
                    min={1}
                    max={10}
                    value={config.walkInDoorQty}
                    onChange={(e) => update("walkInDoorQty", Number(e.target.value))}
                  />
                </div>
              )}
            </div>

            {/* Windows */}
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-2">
                <Label>Window</Label>
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
                  <SelectTrigger><SelectValue placeholder="None" /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="__none__">None</SelectItem>
                    {windows.map((w) => (
                      <SelectItem key={w} value={w}>{w}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              {config.windowType && (
                <div className="space-y-2">
                  <Label>Qty</Label>
                  <Input
                    type="number"
                    min={1}
                    max={20}
                    value={config.windowQty}
                    onChange={(e) => update("windowQty", Number(e.target.value))}
                  />
                </div>
              )}
            </div>

            {/* Roll-Up Doors (Ends) */}
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-2">
                <Label>Roll-Up Door (End)</Label>
                <Select
                  value={config.rollUpEndSize || "__none__"}
                  onValueChange={(v) => {
                    const val = v === "__none__" ? undefined : v;
                    setConfig((prev) => ({
                      ...prev,
                      rollUpEndSize: val,
                      rollUpEndQty: val ? Math.max(1, prev.rollUpEndQty) : 0,
                    }));
                  }}
                >
                  <SelectTrigger><SelectValue placeholder="None" /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="__none__">None</SelectItem>
                    {rollUps.map((r) => (
                      <SelectItem key={r} value={r}>{r}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              {config.rollUpEndSize && (
                <div className="space-y-2">
                  <Label>Qty</Label>
                  <Input
                    type="number"
                    min={1}
                    max={10}
                    value={config.rollUpEndQty}
                    onChange={(e) => update("rollUpEndQty", Number(e.target.value))}
                  />
                </div>
              )}
            </div>

            {/* Roll-Up Doors (Sides) */}
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-2">
                <Label>Roll-Up Door (Side)</Label>
                <Select
                  value={config.rollUpSideSize || "__none__"}
                  onValueChange={(v) => {
                    const val = v === "__none__" ? undefined : v;
                    setConfig((prev) => ({
                      ...prev,
                      rollUpSideSize: val,
                      rollUpSideQty: val ? Math.max(1, prev.rollUpSideQty) : 0,
                    }));
                  }}
                >
                  <SelectTrigger><SelectValue placeholder="None" /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="__none__">None</SelectItem>
                    {rollUps.map((r) => (
                      <SelectItem key={`side-${r}`} value={r}>{r}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              {config.rollUpSideSize && (
                <div className="space-y-2">
                  <Label>Qty</Label>
                  <Input
                    type="number"
                    min={1}
                    max={10}
                    value={config.rollUpSideQty}
                    onChange={(e) => update("rollUpSideQty", Number(e.target.value))}
                  />
                </div>
              )}
            </div>
          </CardContent>
        </Card>

        {/* ── Insulation & Wainscot ── */}
        <Card>
          <CardHeader className="pb-4">
            <CardTitle className="text-base">
              {isWidespan ? "Insulation & Wainscot" : "Insulation"}
            </CardTitle>
          </CardHeader>
          <CardContent className="grid grid-cols-2 gap-4">
            <div className="space-y-2">
              <Label>Insulation</Label>
              <Select
                value={config.insulationType}
                onValueChange={(v) => update("insulationType", v as BuildingConfig["insulationType"])}
              >
                <SelectTrigger><SelectValue /></SelectTrigger>
                <SelectContent>
                  {INSULATION_OPTIONS.map((opt) => (
                    <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            {isWidespan && (
              <div className="space-y-2">
                <Label>Wainscot</Label>
                <Select
                  value={config.wainscot || "none"}
                  onValueChange={(v) => update("wainscot", v as BuildingConfig["wainscot"])}
                >
                  <SelectTrigger><SelectValue /></SelectTrigger>
                  <SelectContent>
                    {WAINSCOT_OPTIONS.map((opt) => (
                      <SelectItem key={opt.value} value={opt.value}>{opt.label}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            )}
          </CardContent>
        </Card>

        {/* ── Engineering & Tax ── */}
        <Card>
          <CardHeader className="pb-4">
            <CardTitle className="text-base">Engineering & Tax</CardTitle>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="grid grid-cols-2 gap-4 sm:grid-cols-4">
              <div className="space-y-2">
                <Label>Snow Load</Label>
                <Select
                  value={config.snowLoad || "__none__"}
                  onValueChange={(v) => update("snowLoad", v === "__none__" ? undefined : v)}
                >
                  <SelectTrigger><SelectValue placeholder="None" /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="__none__">None</SelectItem>
                    {(isWidespan ? WIDESPAN_SNOW_LOAD_OPTIONS : STANDARD_SNOW_LOAD_OPTIONS).map(
                      (opt) => (
                        <SelectItem key={opt.value} value={opt.value}>
                          {opt.label}
                        </SelectItem>
                      )
                    )}
                  </SelectContent>
                </Select>
              </div>

              {!isWidespan && regionStates.length > 0 && (
                <div className="space-y-2">
                  <Label>State</Label>
                  <Select
                    value={config.state || "__none__"}
                    onValueChange={(v) => update("state", v === "__none__" ? undefined : v)}
                  >
                    <SelectTrigger><SelectValue placeholder="Select state" /></SelectTrigger>
                    <SelectContent>
                      <SelectItem value="__none__">None</SelectItem>
                      {regionStates.map((st) => (
                        <SelectItem key={st} value={st}>{st}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              )}

              <div className="space-y-2">
                <Label>Wind Rating (MPH)</Label>
                <Input
                  type="number"
                  min={90}
                  max={180}
                  step={5}
                  value={config.windRating}
                  onChange={(e) => update("windRating", Number(e.target.value))}
                />
              </div>

              <div className="space-y-2">
                <Label>Tax Rate (%)</Label>
                <Input
                  type="number"
                  min={0}
                  max={15}
                  step={0.01}
                  value={Number((config.taxRate * 100).toFixed(4))}
                  onChange={(e) => update("taxRate", Number(e.target.value) / 100)}
                />
              </div>
            </div>

            <div className="flex items-center gap-6">
              <div className="flex items-center gap-2">
                <Switch
                  checked={config.permitRequired ?? false}
                  onCheckedChange={(v) => update("permitRequired", v)}
                />
                <Label>Permit Required</Label>
              </div>

              <div className="flex items-center gap-2">
                <Switch
                  checked={config.diagonalBracing}
                  onCheckedChange={(v) => update("diagonalBracing", v)}
                />
                <Label>Diagonal Bracing</Label>
              </div>

              <div className="flex items-center gap-2">
                <Switch
                  checked={config.includePlans ?? false}
                  onCheckedChange={(v) => update("includePlans", v)}
                />
                <Label>Include Plans</Label>
              </div>
            </div>
          </CardContent>
        </Card>
      </div>

      {/* Right: Price Summary (sticky) */}
      <div className="lg:sticky lg:top-6 lg:self-start space-y-3">
        <Card>
          <CardContent className="pt-6">
            <PriceSummary breakdown={breakdown} isWidespan={isWidespan} />
            {breakdown && (
              <Button
                className="mt-4 w-full"
                onClick={handleSaveQuote}
                disabled={saving}
              >
                {saving ? "Saving..." : "Save Quote"}
              </Button>
            )}
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
