"use client";

import { useCallback, useEffect, useState } from "react";
import Link from "next/link";
import {
  Plus,
  Pencil,
  Trash2,
  Loader2,
  UsersRound,
  Search,
  UserPlus,
  ChevronDown,
  ChevronUp,
} from "lucide-react";
import { AppHeader } from "@/components/layout/app-header";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Switch } from "@/components/ui/switch";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
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
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { formatCurrency } from "@/lib/utils";

type EmployeeRole = "sales_rep" | "sales_manager" | "bst";

interface SalesRep {
  id: string;
  profile_id: string | null;
  name: string;
  email: string;
  phone: string | null;
  office: string;
  territory: string | null;
  commission_rate: number;
  employee_role: EmployeeRole;
  is_active: boolean;
  created_at: string;
  updated_at: string;
  customer_count: number;
  quote_count: number;
  quote_total: number;
}

const ROLE_LABELS: Record<EmployeeRole, string> = {
  sales_rep: "Sales Rep",
  sales_manager: "Sales Manager",
  bst: "BST",
};

const ROLE_COLORS: Record<EmployeeRole, "default" | "secondary" | "destructive"> = {
  sales_rep: "default",
  sales_manager: "secondary",
  bst: "secondary",
};

const ROLE_FILTERS = [
  { value: "all", label: "All Roles" },
  { value: "sales_rep", label: "Sales Reps" },
  { value: "sales_manager", label: "Sales Managers" },
  { value: "bst", label: "BST" },
] as const;

interface PendingProfile {
  id: string;
  email: string;
  name: string | null;
  role: string;
  office: string | null;
  created_at: string;
}

const emptyForm = {
  name: "",
  email: "",
  phone: "",
  office: "Harbor" as string,
  territory: "",
  commission_rate: "0",
  profile_id: "",
  employee_role: "sales_rep" as string,
};

export default function TeamPage() {
  const [reps, setReps] = useState<SalesRep[]>([]);
  const [pending, setPending] = useState<PendingProfile[]>([]);
  const [pendingOpen, setPendingOpen] = useState(true);
  const [loading, setLoading] = useState(true);
  const [toggling, setToggling] = useState<string | null>(null);
  const [dialogOpen, setDialogOpen] = useState(false);
  const [editingRep, setEditingRep] = useState<SalesRep | null>(null);
  const [deleteConfirm, setDeleteConfirm] = useState<SalesRep | null>(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState("");
  const [roleFilter, setRoleFilter] = useState("all");

  const [form, setForm] = useState(emptyForm);

  const set = (field: string, value: string) =>
    setForm((prev) => ({ ...prev, [field]: value }));

  const fetchReps = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch("/api/admin/sales-reps");
      const data = await res.json();
      if (Array.isArray(data)) setReps(data);
    } catch {
      /* ignore */
    } finally {
      setLoading(false);
    }
  }, []);

  const fetchPending = useCallback(async () => {
    try {
      const res = await fetch("/api/admin/sales-reps/pending");
      const data = await res.json();
      if (Array.isArray(data)) setPending(data);
    } catch {
      /* ignore */
    }
  }, []);

  useEffect(() => {
    fetchReps();
    fetchPending();
  }, [fetchReps, fetchPending]);

  const handleToggle = async (rep: SalesRep) => {
    setToggling(rep.id);
    try {
      await fetch(`/api/admin/sales-reps/${rep.id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ is_active: !rep.is_active }),
      });
      await fetchReps();
    } finally {
      setToggling(null);
    }
  };

  const openCreate = () => {
    setEditingRep(null);
    setForm(emptyForm);
    setError(null);
    setDialogOpen(true);
  };

  const openOnboard = (profile: PendingProfile) => {
    setEditingRep(null);
    setForm({
      name: profile.name || profile.email.split("@")[0],
      email: profile.email,
      phone: "",
      office: profile.office || "Harbor",
      territory: "",
      commission_rate: "0",
      profile_id: profile.id,
      employee_role: "sales_rep",
    });
    setError(null);
    setDialogOpen(true);
  };

  const openEdit = (rep: SalesRep) => {
    setEditingRep(rep);
    setForm({
      name: rep.name,
      email: rep.email,
      phone: rep.phone || "",
      office: rep.office,
      territory: rep.territory || "",
      commission_rate: String(rep.commission_rate),
      profile_id: rep.profile_id || "",
      employee_role: rep.employee_role || "sales_rep",
    });
    setError(null);
    setDialogOpen(true);
  };

  const handleSave = async () => {
    if (!form.name.trim()) {
      setError("Name is required");
      return;
    }
    if (!form.email.trim()) {
      setError("Email is required");
      return;
    }

    setSaving(true);
    setError(null);

    try {
      const url = editingRep
        ? `/api/admin/sales-reps/${editingRep.id}`
        : "/api/admin/sales-reps";
      const method = editingRep ? "PATCH" : "POST";

      const payload = {
        name: form.name.trim(),
        email: form.email.trim(),
        phone: form.phone || null,
        office: form.office,
        territory: form.territory || null,
        commission_rate: parseFloat(form.commission_rate) || 0,
        profile_id: form.profile_id || null,
        employee_role: form.employee_role,
      };

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
      await Promise.all([fetchReps(), fetchPending()]);
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!deleteConfirm) return;
    setSaving(true);
    setError(null);

    try {
      const res = await fetch(`/api/admin/sales-reps/${deleteConfirm.id}`, {
        method: "DELETE",
      });
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "Failed to delete");
        return;
      }
      setDeleteConfirm(null);
      await fetchReps();
    } finally {
      setSaving(false);
    }
  };

  const filtered = reps.filter((r) => {
    if (roleFilter !== "all" && r.employee_role !== roleFilter) return false;
    if (!search) return true;
    const s = search.toLowerCase();
    return (
      r.name.toLowerCase().includes(s) ||
      r.email.toLowerCase().includes(s) ||
      (r.territory || "").toLowerCase().includes(s) ||
      r.office.toLowerCase().includes(s)
    );
  });

  const activeCount = reps.filter((r) => r.is_active).length;

  return (
    <>
      <AppHeader title="Team" />
      <div className="flex-1 p-6">
        <div className="mx-auto max-w-6xl space-y-4">
          {/* Header */}
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="relative max-w-sm">
                <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
                <Input
                  placeholder="Search team..."
                  className="pl-9 w-64"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                />
              </div>
              <Select value={roleFilter} onValueChange={setRoleFilter}>
                <SelectTrigger className="w-40">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  {ROLE_FILTERS.map((rf) => (
                    <SelectItem key={rf.value} value={rf.value}>
                      {rf.label}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
              <Badge variant="secondary">
                {activeCount} active / {reps.length} total
              </Badge>
            </div>
            <Button onClick={openCreate}>
              <Plus className="mr-2 h-4 w-4" />
              Add Team Member
            </Button>
          </div>

          {/* Pending Users Banner */}
          {pending.length > 0 && (
            <div className="rounded-lg border border-amber-300 bg-amber-50 dark:border-amber-700 dark:bg-amber-950/30">
              <button
                type="button"
                onClick={() => setPendingOpen((v) => !v)}
                className="w-full flex items-center justify-between px-4 py-3 text-left"
              >
                <div className="flex items-center gap-2">
                  <UserPlus className="h-4 w-4 text-amber-600 dark:text-amber-400" />
                  <span className="text-sm font-medium text-amber-800 dark:text-amber-200">
                    {pending.length} new {pending.length === 1 ? "user has" : "users have"} signed in and {pending.length === 1 ? "needs" : "need"} to be onboarded
                  </span>
                </div>
                {pendingOpen ? (
                  <ChevronUp className="h-4 w-4 text-amber-600 dark:text-amber-400" />
                ) : (
                  <ChevronDown className="h-4 w-4 text-amber-600 dark:text-amber-400" />
                )}
              </button>
              {pendingOpen && (
                <div className="border-t border-amber-300 dark:border-amber-700">
                  <Table>
                    <TableHeader>
                      <TableRow className="hover:bg-transparent">
                        <TableHead>Name</TableHead>
                        <TableHead>Email</TableHead>
                        <TableHead>Signed In</TableHead>
                        <TableHead className="w-28" />
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {pending.map((p) => (
                        <TableRow key={p.id} className="hover:bg-amber-100/50 dark:hover:bg-amber-900/20">
                          <TableCell className="font-medium">
                            {p.name || p.email.split("@")[0]}
                          </TableCell>
                          <TableCell>{p.email}</TableCell>
                          <TableCell className="text-muted-foreground text-xs">
                            {new Intl.DateTimeFormat("en-US", {
                              month: "short",
                              day: "numeric",
                              year: "numeric",
                            }).format(new Date(p.created_at))}
                          </TableCell>
                          <TableCell>
                            <Button
                              size="sm"
                              variant="outline"
                              className="h-7 text-xs"
                              onClick={() => openOnboard(p)}
                            >
                              <UserPlus className="mr-1 h-3 w-3" />
                              Onboard
                            </Button>
                          </TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </div>
              )}
            </div>
          )}

          {loading ? (
            <div className="flex items-center justify-center py-12">
              <Loader2 className="h-8 w-8 animate-spin text-muted-foreground" />
            </div>
          ) : filtered.length === 0 ? (
            <div className="rounded-lg border border-dashed p-12 text-center text-muted-foreground">
              <UsersRound className="mx-auto h-10 w-10 mb-2" />
              <p>{search || roleFilter !== "all" ? "No matching team members." : "No team members yet."}</p>
            </div>
          ) : (
            <div className="rounded-lg border">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="w-12">Active</TableHead>
                    <TableHead>Name</TableHead>
                    <TableHead>Role</TableHead>
                    <TableHead>Email</TableHead>
                    <TableHead>Office</TableHead>
                    <TableHead>Territory</TableHead>
                    <TableHead className="text-right">Commission</TableHead>
                    <TableHead className="text-right">Customers</TableHead>
                    <TableHead className="text-right">Quotes</TableHead>
                    <TableHead className="text-right">Total Sales</TableHead>
                    <TableHead className="w-20" />
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {filtered.map((rep) => (
                    <TableRow
                      key={rep.id}
                      className={!rep.is_active ? "opacity-50" : ""}
                    >
                      <TableCell>
                        {toggling === rep.id ? (
                          <Loader2 className="h-4 w-4 animate-spin" />
                        ) : (
                          <Switch
                            checked={rep.is_active}
                            onCheckedChange={() => handleToggle(rep)}
                            className="scale-90"
                          />
                        )}
                      </TableCell>
                      <TableCell className="font-medium">
                        <Link
                          href={`/admin/sales-reps/${rep.id}`}
                          className="text-primary hover:underline"
                        >
                          {rep.name}
                        </Link>
                      </TableCell>
                      <TableCell>
                        <Badge variant={ROLE_COLORS[rep.employee_role]}>
                          {ROLE_LABELS[rep.employee_role] || rep.employee_role}
                        </Badge>
                      </TableCell>
                      <TableCell>{rep.email}</TableCell>
                      <TableCell>
                        <Badge variant="outline">{rep.office}</Badge>
                      </TableCell>
                      <TableCell>{rep.territory || "---"}</TableCell>
                      <TableCell className="text-right">
                        {rep.commission_rate}%
                      </TableCell>
                      <TableCell className="text-right">
                        {rep.customer_count}
                      </TableCell>
                      <TableCell className="text-right">
                        {rep.quote_count}
                      </TableCell>
                      <TableCell className="text-right">
                        {formatCurrency(rep.quote_total)}
                      </TableCell>
                      <TableCell>
                        <div className="flex gap-1">
                          <Button
                            variant="ghost"
                            size="icon"
                            className="h-7 w-7"
                            onClick={() => openEdit(rep)}
                          >
                            <Pencil className="h-3.5 w-3.5" />
                          </Button>
                          <Button
                            variant="ghost"
                            size="icon"
                            className="h-7 w-7"
                            onClick={() => {
                              setError(null);
                              setDeleteConfirm(rep);
                            }}
                          >
                            <Trash2 className="h-3.5 w-3.5 text-destructive" />
                          </Button>
                        </div>
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          )}
        </div>
      </div>

      {/* Create / Edit Dialog */}
      <Dialog open={dialogOpen} onOpenChange={setDialogOpen}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>
              {editingRep ? "Edit Team Member" : form.profile_id ? "Onboard Team Member" : "Add Team Member"}
            </DialogTitle>
            <DialogDescription>
              {editingRep
                ? "Update team member details."
                : form.profile_id
                  ? "Set up this user as a team member. Their Google sign-in is already linked."
                  : "Add a new team member."}
            </DialogDescription>
          </DialogHeader>
          <div className="space-y-3 py-2">
            <div className="grid grid-cols-2 gap-3">
              <div className="space-y-1">
                <Label>Name *</Label>
                <Input
                  value={form.name}
                  onChange={(e) => set("name", e.target.value)}
                />
              </div>
              <div className="space-y-1">
                <Label>Email *</Label>
                <Input
                  type="email"
                  value={form.email}
                  onChange={(e) => set("email", e.target.value)}
                />
              </div>
              <div className="space-y-1">
                <Label>Role *</Label>
                <Select
                  value={form.employee_role}
                  onValueChange={(v) => set("employee_role", v)}
                >
                  <SelectTrigger>
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="sales_rep">Sales Rep</SelectItem>
                    <SelectItem value="sales_manager">Sales Manager</SelectItem>
                    <SelectItem value="bst">BST (Customer Service)</SelectItem>
                  </SelectContent>
                </Select>
              </div>
              <div className="space-y-1">
                <Label>Phone</Label>
                <Input
                  value={form.phone}
                  onChange={(e) => set("phone", e.target.value)}
                />
              </div>
              <div className="space-y-1">
                <Label>Office *</Label>
                <Select
                  value={form.office}
                  onValueChange={(v) => set("office", v)}
                >
                  <SelectTrigger>
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="Harbor">Harbor</SelectItem>
                    <SelectItem value="Marion">Marion</SelectItem>
                  </SelectContent>
                </Select>
              </div>
              <div className="space-y-1">
                <Label>Territory</Label>
                <Input
                  placeholder="e.g. Ohio / Southeast"
                  value={form.territory}
                  onChange={(e) => set("territory", e.target.value)}
                />
              </div>
              <div className="space-y-1">
                <Label>Commission Rate (%)</Label>
                <Input
                  type="number"
                  min="0"
                  max="100"
                  step="0.5"
                  value={form.commission_rate}
                  onChange={(e) => set("commission_rate", e.target.value)}
                />
              </div>
              <div className="space-y-1">
                <Label>Profile ID (optional)</Label>
                <Input
                  placeholder="UUID — auto-linked on Google sign-in"
                  value={form.profile_id}
                  onChange={(e) => set("profile_id", e.target.value)}
                />
              </div>
            </div>

            {error && <p className="text-sm text-destructive">{error}</p>}
          </div>
          <DialogFooter>
            <Button variant="outline" onClick={() => setDialogOpen(false)}>
              Cancel
            </Button>
            <Button onClick={handleSave} disabled={saving}>
              {saving && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}
              {editingRep ? "Save Changes" : "Add Member"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Delete Confirmation Dialog */}
      <Dialog open={!!deleteConfirm} onOpenChange={() => setDeleteConfirm(null)}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Delete Team Member</DialogTitle>
            <DialogDescription>
              Are you sure you want to permanently delete{" "}
              <strong>{deleteConfirm?.name}</strong>? This action cannot be
              undone.
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
