"use client";

import { useEffect, useState } from "react";
import { Building2, Loader2, Warehouse } from "lucide-react";
import { AppHeader } from "@/components/layout/app-header";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { CalculatorForm } from "@/components/features/calculator/calculator-form";
import type { PricingMatrices, SpreadsheetType } from "@/types/pricing";
import type { AppConfig } from "@/lib/pricing/constants";

interface Region {
  id: string;
  name: string;
  slug: string;
  states: string[];
  spreadsheet_type: string;
}

export default function CalculatorPage() {
  const [spreadsheetType, setSpreadsheetType] = useState<SpreadsheetType | null>(null);
  const [regions, setRegions] = useState<Region[]>([]);
  const [selectedRegion, setSelectedRegion] = useState<string>("");
  const [matrices, setMatrices] = useState<PricingMatrices | null>(null);
  const [loadingMatrices, setLoadingMatrices] = useState(false);
  const [matricesError, setMatricesError] = useState<string | null>(null);
  const [lastUpdated, setLastUpdated] = useState<string | null>(null);
  const [appConfig, setAppConfig] = useState<AppConfig>({});

  // Fetch app config once on mount
  useEffect(() => {
    fetch("/api/admin/config")
      .then((r) => r.json())
      .then((data) => {
        if (data && typeof data === "object" && !data.error) setAppConfig(data);
      })
      .catch(() => {});
  }, []);

  // Fetch regions when type changes
  useEffect(() => {
    if (!spreadsheetType) return;
    fetch(`/api/pricing/regions?type=${spreadsheetType}`)
      .then((r) => r.json())
      .then((data) => {
        if (Array.isArray(data)) setRegions(data);
        setSelectedRegion("");
        setMatrices(null);
      })
      .catch(() => {});
  }, [spreadsheetType]);

  // Fetch pricing matrices when region changes
  useEffect(() => {
    if (!selectedRegion) {
      setMatrices(null);
      setMatricesError(null);
      setLastUpdated(null);
      return;
    }
    setLoadingMatrices(true);
    setMatricesError(null);
    fetch(`/api/pricing/${selectedRegion}`)
      .then(async (r) => {
        if (!r.ok) {
          const body = await r.json().catch(() => ({}));
          throw new Error(body.error || `Failed to load pricing data`);
        }
        return r.json();
      })
      .then((data) => {
        setMatrices(data.matrices as PricingMatrices);
        setLastUpdated(data.created_at || null);
      })
      .catch((err) => {
        setMatricesError(err.message);
        setMatrices(null);
      })
      .finally(() => setLoadingMatrices(false));
  }, [selectedRegion]);

  // Step 1: Choose building type
  if (!spreadsheetType) {
    return (
      <>
        <AppHeader title="Pricing Calculator" />
        <div className="flex-1 p-6">
          <div className="mx-auto max-w-2xl space-y-6">
            <div className="text-center space-y-2">
              <h2 className="text-2xl font-semibold tracking-tight">What type of building?</h2>
              <p className="text-muted-foreground">
                Select the building width range to get started.
              </p>
            </div>
            <div className="grid grid-cols-2 gap-4">
              <Card
                className="cursor-pointer transition-colors hover:border-primary"
                onClick={() => setSpreadsheetType("standard")}
              >
                <CardHeader className="text-center pb-2">
                  <Building2 className="mx-auto h-12 w-12 text-muted-foreground" />
                  <CardTitle className="mt-2">Standard</CardTitle>
                  <CardDescription>12&apos; – 30&apos; wide</CardDescription>
                </CardHeader>
                <CardContent className="text-center text-sm text-muted-foreground">
                  Carports, garages, and small buildings
                </CardContent>
              </Card>
              <Card
                className="cursor-pointer transition-colors hover:border-primary"
                onClick={() => setSpreadsheetType("widespan")}
              >
                <CardHeader className="text-center pb-2">
                  <Warehouse className="mx-auto h-12 w-12 text-muted-foreground" />
                  <CardTitle className="mt-2">Widespan</CardTitle>
                  <CardDescription>32&apos; – 60&apos; wide</CardDescription>
                </CardHeader>
                <CardContent className="text-center text-sm text-muted-foreground">
                  Workshops, barns, and commercial buildings
                </CardContent>
              </Card>
            </div>
          </div>
        </div>
      </>
    );
  }

  // Step 2: Choose region, then configure building
  return (
    <>
      <AppHeader title="Pricing Calculator" />
      <div className="flex-1 p-6">
        <div className="mx-auto max-w-5xl space-y-6">
          {/* Type + Region bar */}
          <div className="flex items-center gap-4">
            <Button
              variant="outline"
              size="sm"
              onClick={() => {
                setSpreadsheetType(null);
                setSelectedRegion("");
                setMatrices(null);
              }}
            >
              &larr; Back
            </Button>
            <div className="flex items-center gap-2 text-sm font-medium">
              {spreadsheetType === "standard" ? (
                <Building2 className="h-4 w-4" />
              ) : (
                <Warehouse className="h-4 w-4" />
              )}
              {spreadsheetType === "standard" ? "Standard (12'–30')" : "Widespan (32'–60')"}
            </div>
            <Select value={selectedRegion} onValueChange={setSelectedRegion}>
              <SelectTrigger className="w-64">
                <SelectValue placeholder="Select a region..." />
              </SelectTrigger>
              <SelectContent>
                {regions.map((region) => (
                  <SelectItem key={region.id} value={region.id}>
                    {region.name}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
            {lastUpdated && (
              <span className="text-sm text-muted-foreground">
                Last Updated{" "}
                {new Date(lastUpdated).toLocaleDateString("en-US", {
                  month: "short",
                  day: "numeric",
                  year: "numeric",
                })}{" "}
                {new Date(lastUpdated).toLocaleTimeString("en-US", {
                  hour: "numeric",
                  minute: "2-digit",
                })}
              </span>
            )}
          </div>

          {/* Calculator content */}
          {!selectedRegion ? (
            <div className="rounded-lg border border-dashed p-12 text-center text-muted-foreground">
              <p>Select a region to start configuring your building.</p>
            </div>
          ) : loadingMatrices ? (
            <div className="rounded-lg border border-dashed p-12 text-center text-muted-foreground">
              <Loader2 className="mx-auto h-8 w-8 animate-spin" />
              <p className="mt-2">Loading pricing data...</p>
            </div>
          ) : matricesError ? (
            <div className="rounded-lg border border-dashed border-destructive p-12 text-center">
              <p className="text-destructive font-medium">{matricesError}</p>
              <p className="mt-1 text-sm text-muted-foreground">
                Upload a spreadsheet for this region in the Admin panel first.
              </p>
            </div>
          ) : (
            <CalculatorForm
              spreadsheetType={spreadsheetType}
              matrices={matrices}
              regionId={selectedRegion}
              regionStates={regions.find((r) => r.id === selectedRegion)?.states || []}
              appConfig={appConfig}
            />
          )}
        </div>
      </div>
    </>
  );
}
