"use client";

import { Suspense, useEffect, useState } from "react";
import { useSearchParams } from "next/navigation";
import { Building2, History, Loader2, Warehouse } from "lucide-react";
import { AppHeader } from "@/components/layout/app-header";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { CalculatorForm } from "@/components/features/calculator/calculator-form";
import type { PricingMatrices, SpreadsheetType } from "@/types/pricing";
import type { AppConfig } from "@/lib/pricing/constants";
import type { BuildingConfig } from "@/types/pricing";

interface Region {
  id: string;
  name: string;
  slug: string;
  states: string[];
  spreadsheet_type: string;
}

interface PricingVersion {
  id: string;
  version: number;
  spreadsheetType: string;
  isCurrent: boolean;
  createdAt: string;
  filename: string | null;
}

interface SourceQuote {
  config: BuildingConfig;
  customer: { name?: string; email?: string; phone?: string; address?: string; city?: string; state?: string; zip?: string };
  notes: string | null;
  region_id: string;
}

export default function CalculatorPageWrapper() {
  return (
    <Suspense>
      <CalculatorPage />
    </Suspense>
  );
}

function CalculatorPage() {
  const searchParams = useSearchParams();
  const fromQuoteId = searchParams.get("from");

  const [spreadsheetType, setSpreadsheetType] = useState<SpreadsheetType | null>(null);
  const [regions, setRegions] = useState<Region[]>([]);
  const [selectedRegion, setSelectedRegion] = useState<string>("");
  const [matrices, setMatrices] = useState<PricingMatrices | null>(null);
  const [loadingMatrices, setLoadingMatrices] = useState(false);
  const [matricesError, setMatricesError] = useState<string | null>(null);
  const [lastUpdated, setLastUpdated] = useState<string | null>(null);
  const [versions, setVersions] = useState<PricingVersion[]>([]);
  const [selectedVersion, setSelectedVersion] = useState<string>("current");
  const [appConfig, setAppConfig] = useState<AppConfig>({});
  const [sourceQuote, setSourceQuote] = useState<SourceQuote | null>(null);
  const [loadingSource, setLoadingSource] = useState(!!fromQuoteId);

  // Load source quote for "edit quote" flow
  useEffect(() => {
    if (!fromQuoteId) return;
    setLoadingSource(true);
    fetch(`/api/quotes/${fromQuoteId}`)
      .then((r) => r.json())
      .then((data) => {
        if (!data.id) return;
        const sq: SourceQuote = {
          config: data.config,
          customer: {
            name: data.customer_name || undefined,
            email: data.customer_email || undefined,
            phone: data.customer_phone || undefined,
            address: data.customer_address || undefined,
            city: data.customer_city || undefined,
            state: data.customer_state || undefined,
            zip: data.customer_zip || undefined,
          },
          notes: data.notes,
          region_id: data.region_id,
        };
        setSourceQuote(sq);
        // Determine spreadsheet type from width
        const w = data.config?.width ?? 0;
        setSpreadsheetType(w >= 32 ? "widespan" : "standard");
      })
      .catch(() => {})
      .finally(() => setLoadingSource(false));
  }, [fromQuoteId]);

  // Fetch app config — re-fetches when region or type changes to get scoped values
  useEffect(() => {
    const params = new URLSearchParams();
    if (selectedRegion) params.set("regionId", selectedRegion);
    if (spreadsheetType) params.set("spreadsheetType", spreadsheetType);
    fetch(`/api/admin/config?${params}`)
      .then((r) => r.json())
      .then((data) => {
        if (data && typeof data === "object" && !data.error) setAppConfig(data);
      })
      .catch(() => {});
  }, [selectedRegion, spreadsheetType]);

  // Fetch regions when type changes
  useEffect(() => {
    if (!spreadsheetType) return;
    fetch(`/api/pricing/regions?type=${spreadsheetType}`)
      .then((r) => r.json())
      .then((data) => {
        if (Array.isArray(data)) {
          setRegions(data);
          // Auto-select region from source quote
          if (sourceQuote && data.some((r: Region) => r.id === sourceQuote.region_id)) {
            setSelectedRegion(sourceQuote.region_id);
          } else {
            setSelectedRegion("");
          }
        }
        setSelectedVersion("current");
        setMatrices(null);
      })
      .catch(() => {});
  }, [spreadsheetType, sourceQuote]);

  // Fetch versions list when region changes
  useEffect(() => {
    if (!selectedRegion) {
      setVersions([]);
      setSelectedVersion("current");
      return;
    }
    fetch(`/api/pricing/${selectedRegion}/versions`)
      .then((r) => r.json())
      .then((data) => {
        if (Array.isArray(data)) setVersions(data);
      })
      .catch(() => {});
  }, [selectedRegion]);

  // Fetch pricing matrices when region or version changes
  useEffect(() => {
    if (!selectedRegion) {
      setMatrices(null);
      setMatricesError(null);
      setLastUpdated(null);
      return;
    }
    setLoadingMatrices(true);
    setMatricesError(null);
    const url =
      selectedVersion === "current"
        ? `/api/pricing/${selectedRegion}`
        : `/api/pricing/${selectedRegion}?versionId=${selectedVersion}`;
    fetch(url)
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
  }, [selectedRegion, selectedVersion]);

  // Loading source quote
  if (loadingSource) {
    return (
      <>
        <AppHeader title="Pricing Calculator" />
        <div className="flex-1 p-6 text-center text-muted-foreground">
          <Loader2 className="mx-auto h-8 w-8 animate-spin" />
          <p className="mt-2">Loading quote...</p>
        </div>
      </>
    );
  }

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
                setSelectedVersion("current");
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
            {selectedRegion && versions.length > 1 && (
              <Select value={selectedVersion} onValueChange={setSelectedVersion}>
                <SelectTrigger className="w-72">
                  <History className="mr-1 h-4 w-4 shrink-0" />
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="current">Current Pricing</SelectItem>
                  {versions
                    .filter((v) => !v.isCurrent)
                    .map((v) => (
                      <SelectItem key={v.id} value={v.id}>
                        v{v.version} &mdash;{" "}
                        {new Date(v.createdAt).toLocaleDateString("en-US", {
                          month: "short",
                          day: "numeric",
                          year: "numeric",
                        })}
                        {v.filename ? ` (${v.filename})` : ""}
                      </SelectItem>
                    ))}
                </SelectContent>
              </Select>
            )}
            {lastUpdated && (
              <span className="text-sm text-muted-foreground">
                {selectedVersion !== "current" && (
                  <span className="mr-1 text-amber-600 font-medium">Viewing old pricing &mdash;</span>
                )}
                Uploaded{" "}
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
              initialConfig={sourceQuote?.config}
              initialCustomer={sourceQuote?.customer}
              initialNotes={sourceQuote?.notes ?? undefined}
            />
          )}
        </div>
      </div>
    </>
  );
}
