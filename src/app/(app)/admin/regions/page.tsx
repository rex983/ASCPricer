"use client";

import { useCallback, useEffect, useState } from "react";
import {
  Plus,
  Building2,
  Warehouse,
  Pencil,
  Trash2,
  Loader2,
  CheckCircle2,
  XCircle,
} from "lucide-react";
import { AppHeader } from "@/components/layout/app-header";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Switch } from "@/components/ui/switch";
import { Input } from "@/components/ui/input";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";

const US_STATES = [
  "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN",
  "IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV",
  "NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN",
  "TX","UT","VT","VA","WA","WV","WI","WY",
];

interface RegionRow {
  id: string;
  name: string;
  slug: string;
  states: string[];
  spreadsheet_type: "standard" | "widespan";
  is_active: boolean;
  created_at: string;
  updated_at: string;
  currentPricing: { version: number; uploadedAt: string } | null;
  lastUpload: { filename: string; uploadedAt: string } | null;
}

export default function RegionsPage() {
  const [regions, setRegions] = useState<RegionRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [toggling, setToggling] = useState<string | null>(null);
  const [dialogOpen, setDialogOpen] = useState(false);
  const [editingRegion, setEditingRegion] = useState<RegionRow | null>(null);
  const [deleteConfirm, setDeleteConfirm] = useState<RegionRow | null>(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Form state
  const [formName, setFormName] = useState("");
  const [formType, setFormType] = useState<"standard" | "widespan">("standard");
  const [formStates, setFormStates] = useState<string[]>([]);
  const [stateInput, setStateInput] = useState("");

  const fetchRegions = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch("/api/admin/regions");
      const data = await res.json();
      if (Array.isArray(data)) setRegions(data);
    } catch {
      /* ignore */
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchRegions();
  }, [fetchRegions]);

  const handleToggle = async (region: RegionRow) => {
    setToggling(region.id);
    try {
      await fetch(`/api/admin/regions/${region.id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ is_active: !region.is_active }),
      });
      await fetchRegions();
    } finally {
      setToggling(null);
    }
  };

  const openCreate = () => {
    setEditingRegion(null);
    setFormName("");
    setFormType("standard");
    setFormStates([]);
    setStateInput("");
    setError(null);
    setDialogOpen(true);
  };

  const openEdit = (region: RegionRow) => {
    setEditingRegion(region);
    setFormName(region.name);
    setFormType(region.spreadsheet_type);
    setFormStates([...region.states]);
    setStateInput("");
    setError(null);
    setDialogOpen(true);
  };

  const addState = (code: string) => {
    const upper = code.toUpperCase().trim();
    if (upper && US_STATES.includes(upper) && !formStates.includes(upper)) {
      setFormStates((prev) => [...prev, upper]);
    }
    setStateInput("");
  };

  const removeState = (code: string) => {
    setFormStates((prev) => prev.filter((s) => s !== code));
  };

  const handleSave = async () => {
    if (!formName.trim()) {
      setError("Name is required");
      return;
    }
    if (formStates.length === 0) {
      setError("Add at least one state");
      return;
    }

    setSaving(true);
    setError(null);

    try {
      const url = editingRegion
        ? `/api/admin/regions/${editingRegion.id}`
        : "/api/admin/regions";
      const method = editingRegion ? "PATCH" : "POST";

      const payload = editingRegion
        ? { name: formName.trim(), states: formStates }
        : { name: formName.trim(), states: formStates, spreadsheet_type: formType };

      const res = await fetch(url, {
        method,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "Failed to save");
        return;
      }

      setDialogOpen(false);
      await fetchRegions();
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!deleteConfirm) return;
    setSaving(true);
    setError(null);

    try {
      const res = await fetch(`/api/admin/regions/${deleteConfirm.id}`, {
        method: "DELETE",
      });
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "Failed to delete");
        return;
      }
      setDeleteConfirm(null);
      await fetchRegions();
    } finally {
      setSaving(false);
    }
  };

  // Group regions by states for display
  const standardRegions = regions.filter((r) => r.spreadsheet_type === "standard");
  const widespanRegions = regions.filter((r) => r.spreadsheet_type === "widespan");

  return (
    <>
      <AppHeader title="Regions" />
      <div className="flex-1 p-6">
        <div className="mx-auto max-w-4xl space-y-6">
          {/* Header */}
          <div className="flex items-center justify-between">
            <div>
              <h2 className="text-lg font-semibold">Region Management</h2>
              <p className="text-sm text-muted-foreground">
                Add, edit, and toggle regions. Each region+size combo is a separate entry.
              </p>
            </div>
            <Button onClick={openCreate}>
              <Plus className="mr-2 h-4 w-4" />
              Add Region
            </Button>
          </div>

          {loading ? (
            <div className="flex items-center justify-center py-12">
              <Loader2 className="h-8 w-8 animate-spin text-muted-foreground" />
            </div>
          ) : (
            <>
              {/* Standard Regions */}
              <div className="space-y-3">
                <div className="flex items-center gap-2 text-sm font-medium text-muted-foreground">
                  <Building2 className="h-4 w-4" />
                  Standard (12&apos;–30&apos;)
                  <Badge variant="secondary">{standardRegions.length}</Badge>
                </div>
                {standardRegions.length === 0 ? (
                  <div className="rounded-lg border border-dashed p-6 text-center text-sm text-muted-foreground">
                    No standard regions yet
                  </div>
                ) : (
                  <div className="grid gap-3">
                    {standardRegions.map((r) => (
                      <RegionCard
                        key={r.id}
                        region={r}
                        toggling={toggling === r.id}
                        onToggle={() => handleToggle(r)}
                        onEdit={() => openEdit(r)}
                        onDelete={() => { setError(null); setDeleteConfirm(r); }}
                      />
                    ))}
                  </div>
                )}
              </div>

              {/* Widespan Regions */}
              <div className="space-y-3">
                <div className="flex items-center gap-2 text-sm font-medium text-muted-foreground">
                  <Warehouse className="h-4 w-4" />
                  Widespan (32&apos;–60&apos;)
                  <Badge variant="secondary">{widespanRegions.length}</Badge>
                </div>
                {widespanRegions.length === 0 ? (
                  <div className="rounded-lg border border-dashed p-6 text-center text-sm text-muted-foreground">
                    No widespan regions yet
                  </div>
                ) : (
                  <div className="grid gap-3">
                    {widespanRegions.map((r) => (
                      <RegionCard
                        key={r.id}
                        region={r}
                        toggling={toggling === r.id}
                        onToggle={() => handleToggle(r)}
                        onEdit={() => openEdit(r)}
                        onDelete={() => { setError(null); setDeleteConfirm(r); }}
                      />
                    ))}
                  </div>
                )}
              </div>
            </>
          )}
        </div>
      </div>

      {/* Create / Edit Dialog */}
      <Dialog open={dialogOpen} onOpenChange={setDialogOpen}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>{editingRegion ? "Edit Region" : "Add Region"}</DialogTitle>
            <DialogDescription>
              {editingRegion
                ? "Update the region name or states."
                : "Create a new pricing region. Each region+size combo is separate."}
            </DialogDescription>
          </DialogHeader>
          <div className="space-y-4 py-2">
            {/* Name */}
            <div className="space-y-1.5">
              <label className="text-sm font-medium">Region Name</label>
              <Input
                placeholder="e.g. Arizona / Colorado / Utah"
                value={formName}
                onChange={(e) => setFormName(e.target.value)}
              />
            </div>

            {/* Type (only for new) */}
            {!editingRegion && (
              <div className="space-y-1.5">
                <label className="text-sm font-medium">Building Size</label>
                <Select
                  value={formType}
                  onValueChange={(v) => setFormType(v as "standard" | "widespan")}
                >
                  <SelectTrigger>
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="standard">Standard (12&apos;–30&apos;)</SelectItem>
                    <SelectItem value="widespan">Widespan (32&apos;–60&apos;)</SelectItem>
                  </SelectContent>
                </Select>
              </div>
            )}

            {/* States */}
            <div className="space-y-1.5">
              <label className="text-sm font-medium">States</label>
              <div className="flex flex-wrap gap-1.5 min-h-[32px]">
                {formStates.map((s) => (
                  <Badge
                    key={s}
                    variant="secondary"
                    className="cursor-pointer hover:bg-destructive hover:text-destructive-foreground"
                    onClick={() => removeState(s)}
                  >
                    {s} &times;
                  </Badge>
                ))}
              </div>
              <div className="flex gap-2">
                <Select
                  value={stateInput}
                  onValueChange={(v) => addState(v)}
                >
                  <SelectTrigger className="flex-1">
                    <SelectValue placeholder="Add a state..." />
                  </SelectTrigger>
                  <SelectContent className="max-h-60">
                    {US_STATES.filter((s) => !formStates.includes(s)).map((s) => (
                      <SelectItem key={s} value={s}>
                        {s}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </div>

            {error && (
              <p className="text-sm text-destructive">{error}</p>
            )}
          </div>
          <DialogFooter>
            <Button variant="outline" onClick={() => setDialogOpen(false)}>
              Cancel
            </Button>
            <Button onClick={handleSave} disabled={saving}>
              {saving && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}
              {editingRegion ? "Save Changes" : "Create Region"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Delete Confirmation Dialog */}
      <Dialog open={!!deleteConfirm} onOpenChange={() => setDeleteConfirm(null)}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Delete Region</DialogTitle>
            <DialogDescription>
              Are you sure you want to permanently delete{" "}
              <strong>{deleteConfirm?.name}</strong> ({deleteConfirm?.spreadsheet_type})?
              This will also remove all associated pricing data and uploads.
              This action cannot be undone.
            </DialogDescription>
          </DialogHeader>
          {error && <p className="text-sm text-destructive">{error}</p>}
          <DialogFooter>
            <Button variant="outline" onClick={() => setDeleteConfirm(null)}>
              Cancel
            </Button>
            <Button
              variant="destructive"
              onClick={handleDelete}
              disabled={saving}
            >
              {saving && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}
              Delete Permanently
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </>
  );
}

function RegionCard({
  region,
  toggling,
  onToggle,
  onEdit,
  onDelete,
}: {
  region: RegionRow;
  toggling: boolean;
  onToggle: () => void;
  onEdit: () => void;
  onDelete: () => void;
}) {
  return (
    <Card className={!region.is_active ? "opacity-60" : undefined}>
      <CardContent className="flex items-center gap-4 py-3 px-4">
        {/* Toggle */}
        <div className="flex items-center gap-2">
          {toggling ? (
            <Loader2 className="h-4 w-4 animate-spin" />
          ) : (
            <Switch
              checked={region.is_active}
              onCheckedChange={onToggle}
              aria-label={`Toggle ${region.name}`}
            />
          )}
        </div>

        {/* Info */}
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2">
            <span className="font-medium truncate">{region.name}</span>
            {!region.is_active && (
              <Badge variant="outline" className="text-muted-foreground">
                Inactive
              </Badge>
            )}
          </div>
          <div className="flex items-center gap-2 mt-0.5">
            <div className="flex flex-wrap gap-1">
              {region.states.map((s) => (
                <Badge key={s} variant="secondary" className="text-xs">
                  {s}
                </Badge>
              ))}
            </div>
          </div>
        </div>

        {/* Pricing status */}
        <div className="text-right text-xs text-muted-foreground shrink-0">
          {region.currentPricing ? (
            <div className="flex items-center gap-1">
              <CheckCircle2 className="h-3.5 w-3.5 text-green-600" />
              <span>
                v{region.currentPricing.version} &mdash;{" "}
                {new Date(region.currentPricing.uploadedAt).toLocaleDateString("en-US", {
                  month: "short",
                  day: "numeric",
                })}
              </span>
            </div>
          ) : (
            <div className="flex items-center gap-1">
              <XCircle className="h-3.5 w-3.5 text-muted-foreground" />
              <span>No pricing</span>
            </div>
          )}
        </div>

        {/* Actions */}
        <div className="flex items-center gap-1 shrink-0">
          <Button variant="ghost" size="icon" onClick={onEdit}>
            <Pencil className="h-4 w-4" />
          </Button>
          <Button variant="ghost" size="icon" onClick={onDelete}>
            <Trash2 className="h-4 w-4 text-destructive" />
          </Button>
        </div>
      </CardContent>
    </Card>
  );
}
