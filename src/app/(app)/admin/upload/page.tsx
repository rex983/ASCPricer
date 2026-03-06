"use client";

import { useCallback, useEffect, useState } from "react";
import { useDropzone } from "react-dropzone";
import { Upload, FileSpreadsheet, CheckCircle2, XCircle, Loader2, AlertTriangle } from "lucide-react";
import { AppHeader } from "@/components/layout/app-header";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Badge } from "@/components/ui/badge";

interface Region {
  id: string;
  name: string;
  slug: string;
  states: string[];
}

interface UploadResult {
  success?: boolean;
  error?: string;
  errors?: string[];
  upload?: {
    id: string;
    filename: string;
    type: string;
    states: string[];
    sheetCount: number;
    version: number;
  };
  validation?: {
    warnings: string[];
  };
}

export default function UploadPage() {
  const [spreadsheetType, setSpreadsheetType] = useState<"standard" | "widespan">("standard");
  const [regions, setRegions] = useState<Region[]>([]);
  const [selectedRegion, setSelectedRegion] = useState<string>("");
  const [uploading, setUploading] = useState(false);
  const [result, setResult] = useState<UploadResult | null>(null);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);

  useEffect(() => {
    fetch(`/api/pricing/regions?type=${spreadsheetType}`)
      .then((r) => r.json())
      .then((data) => {
        if (Array.isArray(data)) setRegions(data);
        setSelectedRegion("");
      })
      .catch(() => {});
  }, [spreadsheetType]);

  const onDrop = useCallback((acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      setSelectedFile(acceptedFiles[0]);
      setResult(null);
    }
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      "application/vnd.ms-excel": [".xls"],
    },
    maxFiles: 1,
  });

  const handleUpload = async () => {
    if (!selectedFile || !selectedRegion) return;

    setUploading(true);
    setResult(null);

    const formData = new FormData();
    formData.append("file", selectedFile);
    formData.append("regionId", selectedRegion);

    try {
      const response = await fetch("/api/admin/upload", {
        method: "POST",
        body: formData,
      });
      const data = await response.json();
      setResult(data);
      if (data.success) {
        setSelectedFile(null);
      }
    } catch {
      setResult({ error: "Network error — please try again" });
    } finally {
      setUploading(false);
    }
  };

  return (
    <>
      <AppHeader title="Upload Pricing" />
      <div className="flex-1 p-6 max-w-2xl">
        <Card>
          <CardHeader>
            <CardTitle>Upload Pricing Spreadsheet</CardTitle>
            <CardDescription>
              Upload an American Steel Carports Excel spreadsheet to update pricing data.
              The system will auto-detect whether it&apos;s a standard or widespan spreadsheet.
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-6">
            {/* Spreadsheet type selector */}
            <div className="space-y-2">
              <label className="text-sm font-medium">Spreadsheet Type</label>
              <div className="flex gap-2">
                <Button
                  variant={spreadsheetType === "standard" ? "default" : "outline"}
                  className="flex-1"
                  onClick={() => setSpreadsheetType("standard")}
                >
                  Standard (12&apos;–30&apos;)
                </Button>
                <Button
                  variant={spreadsheetType === "widespan" ? "default" : "outline"}
                  className="flex-1"
                  onClick={() => setSpreadsheetType("widespan")}
                >
                  Widespan (32&apos;–60&apos;)
                </Button>
              </div>
            </div>

            {/* Region selector */}
            <div className="space-y-2">
              <label className="text-sm font-medium">Region</label>
              <Select value={selectedRegion} onValueChange={setSelectedRegion}>
                <SelectTrigger>
                  <SelectValue placeholder="Select a region..." />
                </SelectTrigger>
                <SelectContent>
                  {regions.map((region) => (
                    <SelectItem key={region.id} value={region.id}>
                      {region.name}
                      {region.states.length > 0 && (
                        <span className="ml-2 text-muted-foreground">
                          ({region.states.join(", ")})
                        </span>
                      )}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
              {regions.length === 0 && (
                <p className="text-sm text-muted-foreground">
                  No regions found. Create regions in the Regions page first.
                </p>
              )}
            </div>

            {/* Dropzone */}
            <div
              {...getRootProps()}
              className={`cursor-pointer rounded-lg border-2 border-dashed p-8 text-center transition-colors ${
                isDragActive
                  ? "border-primary bg-primary/5"
                  : selectedFile
                    ? "border-green-500 bg-green-500/5"
                    : "border-muted-foreground/25 hover:border-muted-foreground/50"
              }`}
            >
              <input {...getInputProps()} />
              {selectedFile ? (
                <div className="flex flex-col items-center gap-2">
                  <FileSpreadsheet className="h-10 w-10 text-green-600" />
                  <p className="font-medium">{selectedFile.name}</p>
                  <p className="text-sm text-muted-foreground">
                    {(selectedFile.size / 1024).toFixed(0)} KB — Click or drag to replace
                  </p>
                </div>
              ) : (
                <div className="flex flex-col items-center gap-2">
                  <Upload className="h-10 w-10 text-muted-foreground" />
                  {isDragActive ? (
                    <p className="font-medium">Drop the spreadsheet here</p>
                  ) : (
                    <>
                      <p className="font-medium">
                        Drag & drop an Excel file here, or click to browse
                      </p>
                      <p className="text-sm text-muted-foreground">
                        Accepts .xlsx files (standard or widespan pricing)
                      </p>
                    </>
                  )}
                </div>
              )}
            </div>

            {/* Upload button */}
            <Button
              onClick={handleUpload}
              disabled={!selectedFile || !selectedRegion || uploading}
              className="w-full"
            >
              {uploading ? (
                <>
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  Parsing & uploading...
                </>
              ) : (
                <>
                  <Upload className="mr-2 h-4 w-4" />
                  Upload & Parse
                </>
              )}
            </Button>

            {/* Result */}
            {result && (
              <div
                className={`rounded-lg border p-4 ${
                  result.success
                    ? "border-green-500/50 bg-green-500/5"
                    : "border-red-500/50 bg-red-500/5"
                }`}
              >
                {result.success && result.upload ? (
                  <div className="space-y-2">
                    <div className="flex items-center gap-2">
                      <CheckCircle2 className="h-5 w-5 text-green-600" />
                      <span className="font-medium text-green-700">Upload successful</span>
                    </div>
                    <div className="text-sm space-y-1">
                      <p>
                        <span className="text-muted-foreground">Type:</span>{" "}
                        <Badge variant="secondary">{result.upload.type}</Badge>
                      </p>
                      <p>
                        <span className="text-muted-foreground">Sheets parsed:</span>{" "}
                        {result.upload.sheetCount}
                      </p>
                      <p>
                        <span className="text-muted-foreground">Version:</span>{" "}
                        {result.upload.version}
                      </p>
                      {result.upload.states.length > 0 && (
                        <p>
                          <span className="text-muted-foreground">States:</span>{" "}
                          {result.upload.states.join(", ")}
                        </p>
                      )}
                    </div>
                    {result.validation?.warnings &&
                      result.validation.warnings.length > 0 && (
                        <div className="mt-2 space-y-1">
                          {result.validation.warnings.map((w, i) => (
                            <div key={i} className="flex items-start gap-2 text-sm text-amber-700">
                              <AlertTriangle className="mt-0.5 h-4 w-4 shrink-0" />
                              {w}
                            </div>
                          ))}
                        </div>
                      )}
                  </div>
                ) : (
                  <div className="space-y-2">
                    <div className="flex items-center gap-2">
                      <XCircle className="h-5 w-5 text-red-600" />
                      <span className="font-medium text-red-700">
                        {result.error || "Upload failed"}
                      </span>
                    </div>
                    {result.errors && (
                      <ul className="text-sm text-red-600 list-disc list-inside">
                        {result.errors.map((e, i) => (
                          <li key={i}>{e}</li>
                        ))}
                      </ul>
                    )}
                  </div>
                )}
              </div>
            )}
          </CardContent>
        </Card>
      </div>
    </>
  );
}
