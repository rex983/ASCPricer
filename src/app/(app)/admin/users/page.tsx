"use client";

import { useCallback, useEffect, useState } from "react";
import { useSession } from "next-auth/react";
import { useRouter } from "next/navigation";
import {
  Loader2,
  Shield,
  Search,
  ShieldCheck,
  ShieldAlert,
  Download,
  Eye,
  ExternalLink,
} from "lucide-react";
import { AppHeader } from "@/components/layout/app-header";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Input } from "@/components/ui/input";
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
import type { KnownRole, UserRole, Office } from "@/types/auth";

interface UserProfile {
  id: string;
  email: string;
  name: string | null;
  role: UserRole;
  office: Office | null;
  is_active: boolean | null;
  created_at: string;
  quote_count: number;
  quote_total: number;
}

const ROLE_LABELS: Record<KnownRole, string> = {
  admin: "Admin",
  senior_manager: "Senior Manager",
  sales_rep: "Sales Rep",
  bst: "BST",
};

const ROLE_COLORS: Record<KnownRole, "default" | "secondary" | "destructive" | "outline"> = {
  admin: "destructive",
  senior_manager: "default",
  sales_rep: "secondary",
  bst: "outline",
};

const ROLE_ICONS: Record<KnownRole, typeof Shield> = {
  admin: ShieldAlert,
  senior_manager: ShieldCheck,
  sales_rep: Shield,
  bst: Shield,
};

const ROLE_FILTERS = [
  { value: "all", label: "All Roles" },
  { value: "admin", label: "Admins" },
  { value: "senior_manager", label: "Senior Managers" },
  { value: "sales_rep", label: "Sales Reps" },
  { value: "bst", label: "BST" },
] as const;

const OFFICE_FILTERS = [
  { value: "all", label: "All Offices" },
  { value: "Sales", label: "Sales" },
  { value: "none", label: "No Office" },
] as const;

// Any launcher role this app doesn't name (junior_manager, team_lead, custom
// roles, ...) still renders, with a neutral badge.
const roleLabel = (r: UserRole) => ROLE_LABELS[r as KnownRole] ?? r.replace(/_/g, " ");
const roleColor = (r: UserRole) => ROLE_COLORS[r as KnownRole] ?? "outline";
const roleIcon = (r: UserRole) => ROLE_ICONS[r as KnownRole] ?? Shield;

export default function UsersPage() {
  const { data: session, status } = useSession();
  const router = useRouter();
  const role = session?.user?.role;
  const isAdminOrManager = role === "admin" || role === "senior_manager";
  const canImpersonate = isAdminOrManager;

  const [users, setUsers] = useState<UserProfile[]>([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [roleFilter, setRoleFilter] = useState("all");
  const [officeFilter, setOfficeFilter] = useState("all");
  const [viewingAs, setViewingAs] = useState<string | null>(null);
  const [launcherUrl, setLauncherUrl] = useState<string | null>(null);

  const fetchUsers = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch("/api/admin/users?limit=200");
      const data = await res.json();
      if (data.users && Array.isArray(data.users)) setUsers(data.users);
      if (typeof data.launcherUrl === "string") setLauncherUrl(data.launcherUrl);
    } catch {
      /* ignore */
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchUsers();
  }, [fetchUsers]);

  useEffect(() => {
    if (status === "authenticated" && !isAdminOrManager) {
      router.replace("/dashboard");
    }
  }, [status, isAdminOrManager, router]);

  const handleViewAs = async (user: UserProfile) => {
    setViewingAs(user.id);
    try {
      const res = await fetch("/api/view-as", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ targetProfileId: user.id }),
      });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        alert(data.error || "Failed to start view-as");
        setViewingAs(null);
        return;
      }
      window.location.href = "/dashboard";
    } catch {
      alert("Failed to start view-as");
      setViewingAs(null);
    }
  };

  const downloadCsv = () => {
    const escape = (v: string) => (v.includes(",") || v.includes('"') ? `"${v.replace(/"/g, '""')}"` : v);
    const header = "name,email,role,office,quotes,total_sales";
    const rows = filtered.map((u) =>
      [
        escape(u.name || ""),
        escape(u.email),
        u.role,
        u.office || "",
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

  // Redirect non-admin/manager users
  if (status === "loading" || (status === "authenticated" && !isAdminOrManager)) {
    return null;
  }

  return (
    <>
      <AppHeader title="Users" />
      <div className="flex-1 p-4 sm:p-6">
        <div className="mx-auto max-w-7xl space-y-4">
          <div className="flex flex-col gap-1 rounded-lg border bg-muted/40 px-4 py-3 text-sm sm:flex-row sm:items-center sm:justify-between">
            <span className="text-muted-foreground">
              Users, roles and offices are managed in the BBD Launcher (Admin &rarr; Users).
              Changes there apply here within a minute.
            </span>
            {launcherUrl && (
              <a
                href={`${launcherUrl}/admin/users`}
                target="_blank"
                rel="noreferrer"
                className="inline-flex items-center gap-1 font-medium text-primary hover:underline"
              >
                Managed in the BBD Launcher
                <ExternalLink className="h-3.5 w-3.5" />
              </a>
            )}
          </div>

          {/* Stats */}
          <div className="flex items-center gap-2 flex-wrap">
            <Badge variant="secondary">{users.length} total</Badge>
            {Object.entries(roleCounts).map(([role, count]) => (
              <Badge key={role} variant="outline" className="text-xs">
                {roleLabel(role)}: {count}
              </Badge>
            ))}
          </div>

          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-3">
              <div className="relative w-full sm:w-auto">
                <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
                <Input
                  placeholder="Search users..."
                  className="pl-9 w-full sm:w-64"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                />
              </div>
              <div className="flex gap-2">
                <Select value={roleFilter} onValueChange={setRoleFilter}>
                  <SelectTrigger className="w-[calc(50vw-2rem)] sm:w-36">
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
                  <SelectTrigger className="w-[calc(50vw-2rem)] sm:w-36">
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
            </div>
            {role === "admin" && (
              <Button variant="outline" onClick={downloadCsv}>
                <Download className="mr-2 h-4 w-4" />
                Download CSV
              </Button>
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
            <div className="rounded-lg border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Name</TableHead>
                    <TableHead>Role</TableHead>
                    <TableHead className="hidden sm:table-cell">Office</TableHead>
                    <TableHead className="hidden sm:table-cell text-right">Quotes</TableHead>
                    <TableHead className="hidden md:table-cell text-right">Total Sales</TableHead>
                    <TableHead className="w-20" />
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {filtered.map((user) => {
                    const RoleIcon = roleIcon(user.role);
                    const isSelf = user.id === session?.user?.profileId;

                    return (
                      <TableRow key={user.id}>
                        <TableCell className="font-medium">
                          <div className="flex items-center gap-2">
                            <span>{user.name || user.email.split("@")[0]}</span>
                            {isSelf && (
                              <Badge variant="outline" className="text-[10px] px-1.5 py-0">
                                You
                              </Badge>
                            )}
                            {user.is_active === false && (
                              <Badge variant="secondary" className="text-[10px] px-1.5 py-0">
                                Inactive
                              </Badge>
                            )}
                          </div>
                          <div className="text-xs text-muted-foreground">{user.email}</div>
                        </TableCell>
                        <TableCell>
                          <Badge variant={roleColor(user.role)}>
                            <RoleIcon className="mr-1 h-3 w-3" />
                            {roleLabel(user.role)}
                          </Badge>
                        </TableCell>
                        <TableCell className="hidden sm:table-cell">
                          {user.office ? (
                            <Badge variant="outline">{user.office}</Badge>
                          ) : (
                            <span className="text-muted-foreground text-xs">---</span>
                          )}
                        </TableCell>
                        <TableCell className="hidden sm:table-cell text-right">
                          {user.quote_count || <span className="text-muted-foreground text-xs">0</span>}
                        </TableCell>
                        <TableCell className="hidden md:table-cell text-right">
                          {user.quote_total > 0 ? formatCurrency(user.quote_total) : <span className="text-muted-foreground text-xs">$0</span>}
                        </TableCell>
                        <TableCell>
                          <div className="flex gap-1">
                            {canImpersonate && !isSelf && user.is_active !== false && ["sales_rep", "bst", "senior_manager"].includes(user.role) && (
                              <Button
                                variant="outline"
                                size="sm"
                                className="h-7 gap-1 px-2 text-xs"
                                disabled={viewingAs === user.id}
                                onClick={() => handleViewAs(user)}
                              >
                                {viewingAs === user.id ? (
                                  <Loader2 className="h-3 w-3 animate-spin" />
                                ) : (
                                  <Eye className="h-3 w-3" />
                                )}
                                View As
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

    </>
  );
}
