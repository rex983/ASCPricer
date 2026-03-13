"use client";

import { useCallback, useEffect, useState } from "react";
import { Loader2, Plus, Trash2, CheckCircle2 } from "lucide-react";
import { AppHeader } from "@/components/layout/app-header";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Separator } from "@/components/ui/separator";

/* ---------- Types ---------- */
type ValueLabelItem = { value: string; label: string };
type ConfigMap = Record<string, unknown>;

/* ---------- Section definitions ---------- */
const SECTIONS: {
  title: string;
  keys: string[];
}[] = [
  {
    title: "Dimensions",
    keys: [
      "standard_widths",
      "widespan_widths",
      "height_range_standard",
      "height_range_widespan",
      "gauge_options",
    ],
  },
  {
    title: "Building Options",
    keys: [
      "roof_styles",
      "side_coverage_options",
      "end_types",
      "orientation_options",
      "sheet_metal_options",
    ],
  },
  {
    title: "Insulation & Wainscot",
    keys: ["insulation_types", "insulation_scopes", "wainscot_options"],
  },
  {
    title: "Engineering",
    keys: [
      "standard_snow_load_options",
      "widespan_snow_load_options",
      "plans_snow_surcharges",
      "brace_prices",
    ],
  },
  {
    title: "Disclaimers",
    keys: ["disclaimers"],
  },
];

const FRIENDLY_LABELS: Record<string, string> = {
  standard_widths: "Standard Widths",
  widespan_widths: "Widespan Widths",
  height_range_standard: "Height Range (Standard)",
  height_range_widespan: "Height Range (Widespan)",
  gauge_options: "Gauge Options",
  roof_styles: "Roof Styles",
  side_coverage_options: "Side Coverage Options",
  end_types: "End Types",
  orientation_options: "Orientation Options",
  sheet_metal_options: "Sheet Metal Options (Widespan)",
  insulation_types: "Insulation Types",
  insulation_scopes: "Insulation Scopes",
  wainscot_options: "Wainscot Options",
  standard_snow_load_options: "Snow Load Options (Standard)",
  widespan_snow_load_options: "Snow Load Options (Widespan)",
  plans_snow_surcharges: "Plans / Snow Surcharges",
  brace_prices: "Brace Prices",
  disclaimers: "Disclaimers",
};

/* ---------- Helpers ---------- */

function isValueLabelArray(val: unknown): val is ValueLabelItem[] {
  return (
    Array.isArray(val) &&
    val.length > 0 &&
    typeof val[0] === "object" &&
    val[0] !== null &&
    "value" in val[0] &&
    "label" in val[0]
  );
}

function isNumberArray(val: unknown): val is number[] {
  return (
    Array.isArray(val) && val.length > 0 && typeof val[0] === "number"
  );
}

function isStringArray(val: unknown): val is string[] {
  return (
    Array.isArray(val) && val.length > 0 && typeof val[0] === "string"
  );
}

function isNumericObject(
  val: unknown
): val is Record<string, number> {
  if (typeof val !== "object" || val === null || Array.isArray(val)) return false;
  return Object.values(val as Record<string, unknown>).every(
    (v) => typeof v === "number"
  );
}

/* ---------- Sub-editors ---------- */

function NumberArrayEditor({
  value,
  onChange,
}: {
  value: number[];
  onChange: (v: number[]) => void;
}) {
  const [text, setText] = useState(value.join(", "));

  // Sync external changes
  useEffect(() => {
    setText(value.join(", "));
  }, [value]);

  return (
    <Input
      value={text}
      onChange={(e) => {
        setText(e.target.value);
        const nums = e.target.value
          .split(",")
          .map((s) => s.trim())
          .filter((s) => s !== "")
          .map(Number)
          .filter((n) => !isNaN(n));
        onChange(nums);
      }}
      placeholder="e.g. 12, 14, 16"
    />
  );
}

function ValueLabelArrayEditor({
  value,
  onChange,
}: {
  value: ValueLabelItem[];
  onChange: (v: ValueLabelItem[]) => void;
}) {
  const update = (idx: number, field: "value" | "label", text: string) => {
    const copy = value.map((item) => ({ ...item }));
    copy[idx][field] = text;
    onChange(copy);
  };

  const remove = (idx: number) => {
    onChange(value.filter((_, i) => i !== idx));
  };

  const add = () => {
    onChange([...value, { value: "", label: "" }]);
  };

  return (
    <div className="space-y-2">
      {value.map((item, idx) => (
        <div key={idx} className="flex items-center gap-2">
          <Input
            className="flex-1"
            placeholder="value"
            value={item.value}
            onChange={(e) => update(idx, "value", e.target.value)}
          />
          <Input
            className="flex-1"
            placeholder="label"
            value={item.label}
            onChange={(e) => update(idx, "label", e.target.value)}
          />
          <Button
            type="button"
            variant="ghost"
            size="icon"
            onClick={() => remove(idx)}
          >
            <Trash2 className="h-4 w-4 text-red-500" />
          </Button>
        </div>
      ))}
      <Button type="button" variant="outline" size="sm" onClick={add}>
        <Plus className="mr-1 h-4 w-4" /> Add Row
      </Button>
    </div>
  );
}

function NumericObjectEditor({
  value,
  onChange,
}: {
  value: Record<string, number>;
  onChange: (v: Record<string, number>) => void;
}) {
  const entries = Object.entries(value);

  const update = (key: string, num: number) => {
    onChange({ ...value, [key]: num });
  };

  const removeKey = (key: string) => {
    const copy = { ...value };
    delete copy[key];
    onChange(copy);
  };

  const addEntry = () => {
    const newKey = `new_key_${Date.now()}`;
    onChange({ ...value, [newKey]: 0 });
  };

  return (
    <div className="space-y-2">
      {entries.map(([key, num]) => (
        <div key={key} className="flex items-center gap-2">
          <Input
            className="flex-1"
            value={key}
            onChange={(e) => {
              const copy = { ...value };
              const val = copy[key];
              delete copy[key];
              copy[e.target.value] = val;
              onChange(copy);
            }}
          />
          <Input
            className="w-28"
            type="number"
            value={num}
            onChange={(e) => update(key, Number(e.target.value))}
          />
          <Button
            type="button"
            variant="ghost"
            size="icon"
            onClick={() => removeKey(key)}
          >
            <Trash2 className="h-4 w-4 text-red-500" />
          </Button>
        </div>
      ))}
      <Button type="button" variant="outline" size="sm" onClick={addEntry}>
        <Plus className="mr-1 h-4 w-4" /> Add Entry
      </Button>
    </div>
  );
}

function StringArrayEditor({
  value,
  onChange,
}: {
  value: string[];
  onChange: (v: string[]) => void;
}) {
  const [text, setText] = useState(value.join("\n"));

  useEffect(() => {
    setText(value.join("\n"));
  }, [value]);

  return (
    <textarea
      className="flex min-h-[160px] w-full rounded-md border border-input bg-background px-3 py-2 text-sm ring-offset-background placeholder:text-muted-foreground focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2"
      value={text}
      onChange={(e) => {
        setText(e.target.value);
        const lines = e.target.value
          .split("\n")
          .filter((l) => l.trim() !== "");
        onChange(lines);
      }}
      placeholder="One item per line"
    />
  );
}

/* ---------- Main page ---------- */

export default function VariablesPage() {
  const [config, setConfig] = useState<ConfigMap>({});
  const [draft, setDraft] = useState<ConfigMap>({});
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState<Record<string, boolean>>({});
  const [success, setSuccess] = useState<Record<string, boolean>>({});

  const fetchConfig = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch("/api/admin/config");
      const data = await res.json();
      if (!res.ok) throw new Error(data.error);
      setConfig(data);
      setDraft(structuredClone(data));
    } catch {
      // ignore
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchConfig();
  }, [fetchConfig]);

  const saveKey = async (key: string) => {
    setSaving((s) => ({ ...s, [key]: true }));
    setSuccess((s) => ({ ...s, [key]: false }));
    try {
      const res = await fetch("/api/admin/config", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ key, value: draft[key] }),
      });
      if (!res.ok) {
        const data = await res.json();
        alert("Save failed: " + (data.error || "Unknown error"));
        return;
      }
      setConfig((prev) => ({ ...prev, [key]: structuredClone(draft[key]) }));
      setSuccess((s) => ({ ...s, [key]: true }));
      setTimeout(() => setSuccess((s) => ({ ...s, [key]: false })), 2000);
    } catch {
      alert("Network error saving config");
    } finally {
      setSaving((s) => ({ ...s, [key]: false }));
    }
  };

  const saveSection = async (keys: string[]) => {
    for (const key of keys) {
      if (draft[key] !== undefined) {
        await saveKey(key);
      }
    }
  };

  const updateDraft = (key: string, value: unknown) => {
    setDraft((prev) => ({ ...prev, [key]: value }));
  };

  const renderEditor = (key: string) => {
    const val = draft[key];
    if (val === undefined) {
      return <p className="text-sm text-muted-foreground">Not configured</p>;
    }

    if (isValueLabelArray(val)) {
      return (
        <ValueLabelArrayEditor
          value={val}
          onChange={(v) => updateDraft(key, v)}
        />
      );
    }

    if (isNumberArray(val)) {
      return (
        <NumberArrayEditor
          value={val}
          onChange={(v) => updateDraft(key, v)}
        />
      );
    }

    if (isStringArray(val)) {
      return (
        <StringArrayEditor
          value={val}
          onChange={(v) => updateDraft(key, v)}
        />
      );
    }

    if (isNumericObject(val)) {
      return (
        <NumericObjectEditor
          value={val}
          onChange={(v) => updateDraft(key, v)}
        />
      );
    }

    // Fallback: JSON textarea
    return (
      <textarea
        className="flex min-h-[120px] w-full rounded-md border border-input bg-background px-3 py-2 text-sm font-mono ring-offset-background focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2"
        value={JSON.stringify(val, null, 2)}
        onChange={(e) => {
          try {
            updateDraft(key, JSON.parse(e.target.value));
          } catch {
            // invalid JSON, ignore
          }
        }}
      />
    );
  };

  if (loading) {
    return (
      <>
        <AppHeader title="Variables" />
        <div className="flex items-center justify-center flex-1 p-6">
          <Loader2 className="h-8 w-8 animate-spin text-muted-foreground" />
        </div>
      </>
    );
  }

  return (
    <>
      <AppHeader title="Variables" />
      <div className="flex-1 p-6 max-w-3xl space-y-6">
        {SECTIONS.map((section) => (
          <Card key={section.title}>
            <CardHeader>
              <CardTitle>{section.title}</CardTitle>
            </CardHeader>
            <CardContent className="space-y-6">
              {section.keys.map((key, idx) => (
                <div key={key}>
                  {idx > 0 && <Separator className="mb-4" />}
                  <div className="space-y-2">
                    <div className="flex items-center gap-2">
                      <Label className="text-sm font-medium">
                        {FRIENDLY_LABELS[key] || key}
                      </Label>
                      {success[key] && (
                        <CheckCircle2 className="h-4 w-4 text-green-600" />
                      )}
                    </div>
                    {renderEditor(key)}
                  </div>
                </div>
              ))}
              <div className="flex justify-end pt-2">
                <Button
                  onClick={() => saveSection(section.keys)}
                  disabled={section.keys.some((k) => saving[k])}
                >
                  {section.keys.some((k) => saving[k]) ? (
                    <>
                      <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                      Saving...
                    </>
                  ) : (
                    "Save"
                  )}
                </Button>
              </div>
            </CardContent>
          </Card>
        ))}
      </div>
    </>
  );
}
