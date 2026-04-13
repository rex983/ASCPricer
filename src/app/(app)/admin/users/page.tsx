"use client";

import { useCallback, useEffect, useState } from "react";
import { useSession } from "next-auth/react";
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
import type { UserRole, Office } from "@/types/auth";

interface Profile {
  id: string;
  email: string;
  name: string | null;
  role: UserRole;
  office: Office | null;
  created_at: string;
  has_sales_rep: boolean;
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
};

export default function UsersPage() {
  const { data: session } = useSession();
  const isAdmin = session?.user?.role === "admin";

  const [users, setUsers] = useState<Profile[]>([]);
  const [loading, setLoading] = useState(true);
  const [dialogOpen, setDialogOpen] = useState(false);
  const [editingUser, setEditingUser] = useState<Profile | null>(null);
  const [deleteConfirm, setDeleteConfirm] = useState<Profile | null>(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState("");
  const [roleFilter, setRoleFilter] = useState("all");
  const [officeFilter, setOfficeFilter] = useState("all");

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

  const openCreate = () => {
    setEditingUser(null);
    setForm(emptyForm);
    setError(null);
    setDialogOpen(true);
  };

  const openEdit = (user: Profile) => {
    setEditingUser(user);
    setForm({
      name: user.name || "",
      email: user.email,
      role: user.role,
      office: user.office || "",
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
      };

      // Only admins can set roles; managers editing shouldn't send role
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
      u.email.toLowerCase().includes(s)
    );
  });

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
        <div className="mx-auto max-w-6xl space-y-4">
          {/* Stats */}
          <div className="flex items-center gap-2 flex-wrap">
            <Badge variant="secondary">{users.length} total</Badge>
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
                  CSV import: {importResult.created} created, {importResult.skipped} skipped (already exist){importResult.errors.length > 0 && `, ${importResult.errors.length} failed`}.
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
                    <TableHead>Name</TableHead>
                    <TableHead>Email</TableHead>
                    <TableHead>Role</TableHead>
                    <TableHead>Office</TableHead>
                    <TableHead>Sales Rep</TableHead>
                    <TableHead>Joined</TableHead>
                    <TableHead className="w-20" />
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {filtered.map((user) => {
                    const RoleIcon = ROLE_ICONS[user.role];
                    const isSelf = user.id === session?.user?.profileId;
                    const isPrimaryAdmin = user.email === "rex@bigbuildingsdirect.com";

                    return (
                      <TableRow key={user.id}>
                        <TableCell className="font-medium">
                          <div className="flex items-center gap-2">
                            {user.name || user.email.split("@")[0]}
                            {isSelf && (
                              <Badge variant="outline" className="text-[10px] px-1.5 py-0">
                                You
                              </Badge>
                            )}
                          </div>
                        </TableCell>
                        <TableCell className="text-muted-foreground">
                          {user.email}
                        </TableCell>
                        <TableCell>
                          <Badge variant={ROLE_COLORS[user.role]}>
                            <RoleIcon className="mr-1 h-3 w-3" />
                            {ROLE_LABELS[user.role]}
                          </Badge>
                        </TableCell>
                        <TableCell>
                          {user.office ? (
                            <Badge variant="outline">{user.office}</Badge>
                          ) : (
                            <span className="text-muted-foreground text-xs">---</span>
                          )}
                        </TableCell>
                        <TableCell>
                          {user.has_sales_rep ? (
                            <Badge variant="secondary" className="text-xs">
                              Linked
                            </Badge>
                          ) : (
                            <span className="text-muted-foreground text-xs">---</span>
                          )}
                        </TableCell>
                        <TableCell className="text-muted-foreground text-xs">
                          {formatDate(user.created_at)}
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

      {/* Create / Edit Dialog */}
      <Dialog open={dialogOpen} onOpenChange={setDialogOpen}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>
              {editingUser ? "Edit User" : "Add User"}
            </DialogTitle>
            <DialogDescription>
              {editingUser
                ? "Update this user's profile and permissions."
                : "Create a new user profile. They can sign in with Google using this email."}
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
                  <Label>Role *</Label>
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

      {/* Delete Confirmation Dialog */}
      <Dialog open={!!deleteConfirm} onOpenChange={() => setDeleteConfirm(null)}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Delete User</DialogTitle>
            <DialogDescription>
              Are you sure you want to permanently delete{" "}
              <strong>{deleteConfirm?.name || deleteConfirm?.email}</strong>?
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
