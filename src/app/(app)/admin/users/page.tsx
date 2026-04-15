"use client";

import { useCallback, useEffect, useState } from "react";
import { useSession } from "next-auth/react";
import Link from "next/link";
import {
  Plus,
  Pencil,
  Trash2,
  Loader2,
  Shield,
  Search,
  ShieldCheck,
  ShieldAlert,
  Upload,
  CheckCircle2,
  AlertCircle,
  Download,
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
import type { UserRole, Office } from "@/types/auth";

type EmployeeRole = "sales_rep" | "sales_manager" | "bst";

interface UserProfile {
  id: string;
  email: string;
  name: string | null;
  role: UserRole;
  office: Office | null;
  created_at: string;
  // Sales rep data (null if no linked rep)
  rep_id: string | null;
  phone: string | null;
  territory: string | null;
  commission_rate: number | null;
  employee_role: EmployeeRole | null;
  is_active: boolean | null;
  // Stats
  customer_count: number;
  quote_count: number;
  quote_total: number;
}

const ROLE_LABELS: Record<UserRole, string> = {
  admin: "Admin",
  manager: "Manager",
  sales_rep: "Sales Rep",
  viewer: "Viewer",
};

const ROLE_COLORS: Record<UserRole, "default" | "secondary" | "destructive" | "outline"> = {
  admin: "destructive",
  manager: "default",
  sales_rep: "secondary",
  viewer: "outline",
};

const ROLE_ICONS: Record<UserRole, typeof Shield> = {
  admin: ShieldAlert,
  manager: ShieldCheck,
  sales_rep: Shield,
  viewer: Shield,
};

const EMPLOYEE_ROLE_LABELS: Record<string, string> = {
  sales_rep: "Sales Rep",
  sales_manager: "Sales Manager",
  bst: "BST",
};

const ROLE_FILTERS = [
  { value: "all", label: "All Roles" },
  { value: "admin", label: "Admins" },
  { value: "manager", label: "Managers" },
  { value: "sales_rep", label: "Sales Reps" },
  { value: "viewer", label: "Viewers" },
] as const;

const OFFICE_FILTERS = [
  { value: "all", label: "All Offices" },
  { value: "Harbor", label: "Harbor" },
  { value: "Marion", label: "Marion" },
  { value: "none", label: "No Office" },
] as const;

const emptyForm = {
  name: "",
  email: "",
  role: "viewer" as string,
  office: "" as string,
  phone: "",
  employee_role: "" as string,
  territory: "",
  commission_rate: "0",
};

export default function UsersPage() {
  const { data: session } = useSession();
  const isAdmin = session?.user?.role === "admin";

  const [users, setUsers] = useState<UserProfile[]>([]);
  const [loading, setLoading] = useState(true);
  const [dialogOpen, setDialogOpen] = useState(false);
  const [editingUser, setEditingUser] = useState<UserProfile | null>(null);
  const [deleteConfirm, setDeleteConfirm] = useState<UserProfile | null>(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState("");
  const [roleFilter, setRoleFilter] = useState("all");
  const [officeFilter, setOfficeFilter] = useState("all");
  const [toggling, setToggling] = useState<string | null>(null);

  const [importing, setImporting] = useState(false);
  const [importDialogOpen, setImportDialogOpen] = useState(false);
  const [importResult, setImportResult] = useState<{
    total: number;
    created: number;
    skipped: number;
    errors: { row: number; email: string; reason: string }[];
  } | null>(null);

  const [form, setForm] = useState(emptyForm);

  const set = (field: string, value: string) =>
    setForm((prev) => ({ ...prev, [field]: value }));

  const fetchUsers = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch("/api/admin/users?limit=200");
      const data = await res.json();
      if (data.users && Array.isArray(data.users)) setUsers(data.users);
    } catch {
      /* ignore */
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchUsers();
  }, [fetchUsers]);

  const handleToggle = async (user: UserProfile) => {
    if (!user.rep_id) return;
    setToggling(user.id);
    try {
      await fetch(`/api/admin/users/${user.id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ is_active: !user.is_active }),
      });
      await fetchUsers();
    } finally {
      setToggling(null);
    }
  };

  const handleCsvUpload = async (file: File) => {
    setImporting(true);
    setImportResult(null);
    try {
      const formData = new FormData();
      formData.append("file", file);
      const res = await fetch("/api/admin/users/import", {
        method: "POST",
        body: formData,
      });
      const data = await res.json();
      if (!res.ok) {
        alert(data.error || "Import failed");
        return;
      }
      setImportResult(data);
      setImportDialogOpen(false);
      await fetchUsers();
    } catch {
      alert("Import failed — check console");
    } finally {
      setImporting(false);
    }
  };

  const downloadTemplate = () => {
    const csv = "name,email,role,office\nJohn Doe,john@bigbuildingsdirect.com,viewer,Harbor\n";
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "users-template.csv";
    a.click();
    URL.revokeObjectURL(url);
  };

  const downloadCsv = () => {
    const escape = (v: string) => (v.includes(",") || v.includes('"') ? `"${v.replace(/"/g, '""')}"` : v);
    const header = "name,email,role,office,position,phone,territory,commission_rate,active,customers,quotes,total_sales";
    const rows = filtered.map((u) =>
      [
        escape(u.name || ""),
        escape(u.email),
        u.role,
        u.office || "",
        u.employee_role ? (EMPLOYEE_ROLE_LABELS[u.employee_role] || u.employee_role) : "",
        u.phone || "",
        escape(u.territory || ""),
        u.commission_rate !== null ? String(u.commission_rate) : "",
        u.is_active !== null ? (u.is_active ? "yes" : "no") : "",
        String(u.customer_count),
        String(u.quote_count),
        String(u.quote_total),
      ].join(",")
    );
    const csv = [header, ...rows].join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `users-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const openCreate = () => {
    setEditingUser(null);
    setForm(emptyForm);
    setError(null);
    setDialogOpen(true);
  };

  const openEdit = (user: UserProfile) => {
    setEditingUser(user);
    setForm({
      name: user.name || "",
      email: user.email,
      role: user.role,
      office: user.office || "",
      phone: user.phone || "",
      employee_role: user.employee_role || "",
      territory: user.territory || "",
      commission_rate: String(user.commission_rate ?? 0),
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
      const url = editingUser
        ? `/api/admin/users/${editingUser.id}`
        : "/api/admin/users";
      const method = editingUser ? "PATCH" : "POST";

      const payload: Record<string, unknown> = {
        name: form.name.trim(),
        email: form.email.trim(),
        office: form.office || null,
        phone: form.phone || null,
        territory: form.territory || null,
        commission_rate: parseFloat(form.commission_rate) || 0,
        employee_role: form.employee_role || null,
      };

      if (isAdmin) {
        payload.role = form.role;
      }

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
      await fetchUsers();
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!deleteConfirm) return;
    setSaving(true);
    setError(null);

    try {
      const res = await fetch(`/api/admin/users/${deleteConfirm.id}`, {
        method: "DELETE",
      });
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "Failed to delete");
        return;
      }
      setDeleteConfirm(null);
      await fetchUsers();
    } finally {
      setSaving(false);
    }
  };

  const filtered = users.filter((u) => {
    if (roleFilter !== "all" && u.role !== roleFilter) return false;
    if (officeFilter !== "all") {
      if (officeFilter === "none" && u.office) return false;
      if (officeFilter !== "none" && u.office !== officeFilter) return false;
    }
    if (!search) return true;
    const s = search.toLowerCase();
    return (
      (u.name || "").toLowerCase().includes(s) ||
      u.email.toLowerCase().includes(s) ||
      (u.territory || "").toLowerCase().includes(s)
    );
  });

  const activeCount = users.filter((u) => u.is_active === true).length;
  const roleCounts = users.reduce(
    (acc, u) => {
      acc[u.role] = (acc[u.role] || 0) + 1;
      return acc;
    },
    {} as Record<string, number>
  );

  const formatDate = (date: string) =>
    new Intl.DateTimeFormat("en-US", {
      month: "short",
      day: "numeric",
      year: "numeric",
    }).format(new Date(date));

  return (
    <>
      <AppHeader title="User Management" />
      <div className="flex-1 p-6">
        <div className="mx-auto max-w-7xl space-y-4">
          {/* Stats */}
          <div className="flex items-center gap-2 flex-wrap">
            <Badge variant="secondary">{users.length} total</Badge>
            <Badge variant="secondary">{activeCount} active</Badge>
            {Object.entries(roleCounts).map(([role, count]) => (
              <Badge key={role} variant="outline" className="text-xs">
                {ROLE_LABELS[role as UserRole] || role}: {count}
              </Badge>
            ))}
          </div>

          {/* Import Result Banner */}
          {importResult && (
            <div className={`flex items-start gap-2 rounded-lg border px-4 py-3 ${
              importResult.errors.length > 0
                ? "border-amber-300 bg-amber-50 dark:border-amber-700 dark:bg-amber-950/30"
                : "border-green-300 bg-green-50 dark:border-green-700 dark:bg-green-950/30"
            }`}>
              {importResult.errors.length > 0 ? (
                <AlertCircle className="mt-0.5 h-4 w-4 shrink-0 text-amber-600 dark:text-amber-400" />
              ) : (
                <CheckCircle2 className="mt-0.5 h-4 w-4 shrink-0 text-green-600 dark:text-green-400" />
              )}
              <div className="flex-1 space-y-1">
                <span className={`text-sm ${
                  importResult.errors.length > 0
                    ? "text-amber-800 dark:text-amber-200"
                    : "text-green-800 dark:text-green-200"
                }`}>
                  CSV import: {importResult.created} created, {importResult.skipped} skipped{importResult.errors.length > 0 && `, ${importResult.errors.length} failed`}.
                </span>
                {importResult.errors.length > 0 && (
                  <ul className="text-xs text-amber-700 dark:text-amber-300 space-y-0.5">
                    {importResult.errors.map((e, i) => (
                      <li key={i}>Row {e.row}: {e.email} — {e.reason}</li>
                    ))}
                  </ul>
                )}
              </div>
              <button
                type="button"
                onClick={() => setImportResult(null)}
                className={`text-xs hover:underline ${
                  importResult.errors.length > 0
                    ? "text-amber-600 dark:text-amber-400"
                    : "text-green-600 dark:text-green-400"
                }`}
              >
                Dismiss
              </button>
            </div>
          )}

          {/* Header */}
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="relative max-w-sm">
                <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
                <Input
                  placeholder="Search users..."
                  className="pl-9 w-64"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                />
              </div>
              <Select value={roleFilter} onValueChange={setRoleFilter}>
                <SelectTrigger className="w-36">
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
              <Select value={officeFilter} onValueChange={setOfficeFilter}>
                <SelectTrigger className="w-36">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  {OFFICE_FILTERS.map((of_) => (
                    <SelectItem key={of_.value} value={of_.value}>
                      {of_.label}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            {isAdmin && (
              <div className="flex items-center gap-2">
                <Button variant="outline" onClick={downloadCsv}>
                  <Download className="mr-2 h-4 w-4" />
                  Download CSV
                </Button>
                <Button variant="outline" onClick={() => setImportDialogOpen(true)}>
                  <Upload className="mr-2 h-4 w-4" />
                  Import CSV
                </Button>
                <Button onClick={openCreate}>
                  <Plus className="mr-2 h-4 w-4" />
                  Add User
                </Button>
              </div>
            )}
          </div>

          {loading ? (
            <div className="flex items-center justify-center py-12">
              <Loader2 className="h-8 w-8 animate-spin text-muted-foreground" />
            </div>
          ) : filtered.length === 0 ? (
            <div className="rounded-lg border border-dashed p-12 text-center text-muted-foreground">
              <Shield className="mx-auto h-10 w-10 mb-2" />
              <p>
                {search || roleFilter !== "all" || officeFilter !== "all"
                  ? "No matching users."
                  : "No users yet."}
              </p>
            </div>
          ) : (
            <div className="rounded-lg border">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="w-12">Active</TableHead>
                    <TableHead>Name</TableHead>
                    <TableHead>Role</TableHead>
                    <TableHead>Position</TableHead>
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
                  {filtered.map((user) => {
                    const RoleIcon = ROLE_ICONS[user.role];
                    const isSelf = user.id === session?.user?.profileId;
                    const isPrimaryAdmin = user.email === "rex@bigbuildingsdirect.com";
                    const hasRep = user.rep_id !== null;

                    return (
                      <TableRow
                        key={user.id}
                        className={hasRep && user.is_active === false ? "opacity-50" : ""}
                      >
                        <TableCell>
                          {hasRep ? (
                            toggling === user.id ? (
                              <Loader2 className="h-4 w-4 animate-spin" />
                            ) : (
                              <Switch
                                checked={user.is_active ?? false}
                                onCheckedChange={() => handleToggle(user)}
                                className="scale-90"
                              />
                            )
                          ) : (
                            <span className="text-muted-foreground text-xs">---</span>
                          )}
                        </TableCell>
                        <TableCell className="font-medium">
                          <div className="flex items-center gap-2">
                            {hasRep ? (
                              <Link
                                href={`/admin/sales-reps/${user.rep_id}`}
                                className="text-primary hover:underline"
                              >
                                {user.name || user.email.split("@")[0]}
                              </Link>
                            ) : (
                              <span>{user.name || user.email.split("@")[0]}</span>
                            )}
                            {isSelf && (
                              <Badge variant="outline" className="text-[10px] px-1.5 py-0">
                                You
                              </Badge>
                            )}
                          </div>
                          <div className="text-xs text-muted-foreground">{user.email}</div>
                        </TableCell>
                        <TableCell>
                          <Badge variant={ROLE_COLORS[user.role]}>
                            <RoleIcon className="mr-1 h-3 w-3" />
                            {ROLE_LABELS[user.role]}
                          </Badge>
                        </TableCell>
                        <TableCell>
                          {user.employee_role ? (
                            <Badge variant="secondary" className="text-xs">
                              {EMPLOYEE_ROLE_LABELS[user.employee_role] || user.employee_role}
                            </Badge>
                          ) : (
                            <span className="text-muted-foreground text-xs">---</span>
                          )}
                        </TableCell>
                        <TableCell>
                          {user.office ? (
                            <Badge variant="outline">{user.office}</Badge>
                          ) : (
                            <span className="text-muted-foreground text-xs">---</span>
                          )}
                        </TableCell>
                        <TableCell>
                          {user.territory || <span className="text-muted-foreground text-xs">---</span>}
                        </TableCell>
                        <TableCell className="text-right">
                          {user.commission_rate !== null ? `${user.commission_rate}%` : <span className="text-muted-foreground text-xs">---</span>}
                        </TableCell>
                        <TableCell className="text-right">
                          {user.customer_count || <span className="text-muted-foreground text-xs">0</span>}
                        </TableCell>
                        <TableCell className="text-right">
                          {user.quote_count || <span className="text-muted-foreground text-xs">0</span>}
                        </TableCell>
                        <TableCell className="text-right">
                          {user.quote_total > 0 ? formatCurrency(user.quote_total) : <span className="text-muted-foreground text-xs">$0</span>}
                        </TableCell>
                        <TableCell>
                          <div className="flex gap-1">
                            <Button
                              variant="ghost"
                              size="icon"
                              className="h-7 w-7"
                              onClick={() => openEdit(user)}
                            >
                              <Pencil className="h-3.5 w-3.5" />
                            </Button>
                            {isAdmin && !isSelf && !isPrimaryAdmin && (
                              <Button
                                variant="ghost"
                                size="icon"
                                className="h-7 w-7"
                                onClick={() => {
                                  setError(null);
                                  setDeleteConfirm(user);
                                }}
                              >
                                <Trash2 className="h-3.5 w-3.5 text-destructive" />
                              </Button>
                            )}
                          </div>
                        </TableCell>
                      </TableRow>
                    );
                  })}
                </TableBody>
              </Table>
            </div>
          )}
        </div>
      </div>

      {/* Import CSV Dialog */}
      <Dialog open={importDialogOpen} onOpenChange={setImportDialogOpen}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Import Users from CSV</DialogTitle>
            <DialogDescription>
              Upload a CSV file with columns: <strong>name</strong>, <strong>email</strong> (required), and optionally <strong>role</strong> (admin, manager, sales_rep, viewer) and <strong>office</strong> (Harbor, Marion). Existing emails will be skipped.
            </DialogDescription>
          </DialogHeader>
          <div className="space-y-4 py-2">
            <div className="flex items-center justify-center w-full">
              <label className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed rounded-lg cursor-pointer hover:bg-muted/50 transition-colors">
                <div className="flex flex-col items-center justify-center pt-5 pb-6">
                  {importing ? (
                    <Loader2 className="h-8 w-8 animate-spin text-muted-foreground mb-2" />
                  ) : (
                    <Upload className="h-8 w-8 text-muted-foreground mb-2" />
                  )}
                  <p className="text-sm text-muted-foreground">
                    {importing ? "Importing..." : "Click to upload or drag & drop"}
                  </p>
                  <p className="text-xs text-muted-foreground mt-1">.csv files only</p>
                </div>
                <input
                  type="file"
                  accept=".csv"
                  className="hidden"
                  disabled={importing}
                  onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (file) handleCsvUpload(file);
                    e.target.value = "";
                  }}
                />
              </label>
            </div>
          </div>
          <DialogFooter>
            <Button variant="outline" size="sm" onClick={downloadTemplate}>
              <Download className="mr-2 h-3.5 w-3.5" />
              Download Template
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Create / Edit Dialog */}
      <Dialog open={dialogOpen} onOpenChange={setDialogOpen}>
        <DialogContent className="max-w-lg">
          <DialogHeader>
            <DialogTitle>
              {editingUser ? "Edit User" : "Add User"}
            </DialogTitle>
            <DialogDescription>
              {editingUser
                ? "Update this user's profile and team settings."
                : "Create a new user. Set a position to add them to the sales team."}
            </DialogDescription>
          </DialogHeader>
          <div className="space-y-3 py-2">
            <div className="grid grid-cols-2 gap-3">
              <div className="space-y-1">
                <Label>Name *</Label>
                <Input
                  value={form.name}
                  onChange={(e) => set("name", e.target.value)}
                  placeholder="Full name"
                />
              </div>
              <div className="space-y-1">
                <Label>Email *</Label>
                <Input
                  type="email"
                  value={form.email}
                  onChange={(e) => set("email", e.target.value)}
                  placeholder="user@bigbuildingsdirect.com"
                />
              </div>
              {isAdmin && (
                <div className="space-y-1">
                  <Label>App Role *</Label>
                  <Select
                    value={form.role}
                    onValueChange={(v) => set("role", v)}
                  >
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="admin">Admin</SelectItem>
                      <SelectItem value="manager">Manager</SelectItem>
                      <SelectItem value="sales_rep">Sales Rep</SelectItem>
                      <SelectItem value="viewer">Viewer</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
              )}
              <div className="space-y-1">
                <Label>Office</Label>
                <Select
                  value={form.office || "none"}
                  onValueChange={(v) => set("office", v === "none" ? "" : v)}
                >
                  <SelectTrigger>
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="none">No Office</SelectItem>
                    <SelectItem value="Harbor">Harbor</SelectItem>
                    <SelectItem value="Marion">Marion</SelectItem>
                  </SelectContent>
                </Select>
              </div>
            </div>

            {/* Sales Team Section */}
            <div className="border-t pt-3 mt-3">
              <p className="text-sm font-medium mb-2">Sales Team</p>
              <div className="grid grid-cols-2 gap-3">
                <div className="space-y-1">
                  <Label>Position</Label>
                  <Select
                    value={form.employee_role || "none"}
                    onValueChange={(v) => set("employee_role", v === "none" ? "" : v)}
                  >
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="none">Not on sales team</SelectItem>
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
                    placeholder="(555) 123-4567"
                  />
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
              {editingUser ? "Save Changes" : "Add User"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Delete Confirmation Dialog */}
      <Dialog open={!!deleteConfirm} onOpenChange={() => setDeleteConfirm(null)}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Delete User</DialogTitle>
            <DialogDescription>
              Are you sure you want to permanently delete{" "}
              <strong>{deleteConfirm?.name || deleteConfirm?.email}</strong>?
              {deleteConfirm?.rep_id && " This will also remove their sales rep profile."}
              {" "}This action cannot be undone.
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
