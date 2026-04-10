"use client";

import { useCallback, useEffect, useState } from "react";
import { useParams } from "next/navigation";
import Link from "next/link";
import {
  ArrowLeft,
  DollarSign,
  FileText,
  Loader2,
  Mail,
  MapPin,
  Phone,
  TrendingUp,
  Users,
  Activity,
} from "lucide-react";
import { AppHeader } from "@/components/layout/app-header";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { formatCurrency, formatDate, STATUS_COLORS } from "@/lib/utils";

type EmployeeRole = "sales_rep" | "sales_manager" | "bst";

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

interface RepDetail {
  rep: {
    id: string;
    name: string;
    email: string;
    phone: string | null;
    office: string;
    territory: string | null;
    commission_rate: number;
    employee_role: EmployeeRole;
    is_active: boolean;
    profile_id: string | null;
    created_at: string;
  };
  quotes: {
    id: string;
    quote_number: string;
    status: string;
    customer_name: string | null;
    total: number;
    created_at: string;
  }[];
  customers: {
    id: string;
    name: string;
    email: string | null;
    phone: string | null;
    city: string | null;
    state: string | null;
    office: string;
    created_at: string;
  }[];
  activity: {
    id: string;
    action: string;
    resource_type: string | null;
    resource_id: string | null;
    details: Record<string, unknown>;
    created_at: string;
  }[];
  kpis: {
    totalQuotes: number;
    totalRevenue: number;
    totalCustomers: number;
    acceptedQuotes: number;
    conversionRate: number;
    statusCounts: Record<string, number>;
  };
  monthlyRevenue: {
    month: string;
    revenue: number;
    count: number;
  }[];
}

const ACTION_LABELS: Record<string, string> = {
  create_quote: "Created a quote",
  update_quote: "Updated a quote",
  delete_quote: "Deleted a quote",
  create_customer: "Added a customer",
  update_customer: "Updated a customer",
  delete_customer: "Deleted a customer",
  upload_pricing: "Uploaded pricing data",
  create_sales_rep: "Added a team member",
  update_sales_rep: "Updated a team member",
  toggle_sales_rep: "Toggled team member status",
  update_config: "Updated app config",
};

export default function TeamMemberDetailPage() {
  const params = useParams();
  const id = params.id as string;

  const [data, setData] = useState<RepDetail | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [quoteStatusFilter, setQuoteStatusFilter] = useState("all");

  const fetchDetail = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(`/api/admin/sales-reps/${id}/detail`);
      if (!res.ok) {
        const body = await res.json();
        setError(body.error || "Failed to load");
        return;
      }
      const json = await res.json();
      setData(json);
    } catch {
      setError("Network error");
    } finally {
      setLoading(false);
    }
  }, [id]);

  useEffect(() => {
    fetchDetail();
  }, [fetchDetail]);

  if (loading) {
    return (
      <>
        <AppHeader title="Team Member" />
        <div className="flex-1 flex items-center justify-center py-12">
          <Loader2 className="h-8 w-8 animate-spin text-muted-foreground" />
        </div>
      </>
    );
  }

  if (error || !data) {
    return (
      <>
        <AppHeader title="Team Member" />
        <div className="flex-1 p-6 text-center text-muted-foreground">
          <p>{error || "Not found"}</p>
          <Button asChild variant="outline" className="mt-4">
            <Link href="/admin/sales-reps">Back to Team</Link>
          </Button>
        </div>
      </>
    );
  }

  const { rep, quotes, customers, activity, kpis, monthlyRevenue } = data;

  const filteredQuotes =
    quoteStatusFilter === "all"
      ? quotes
      : quotes.filter((q) => q.status === quoteStatusFilter);

  const maxRevenue = Math.max(...monthlyRevenue.map((m) => m.revenue), 1);

  return (
    <>
      <AppHeader title={rep.name} />
      <div className="flex-1 p-6">
        <div className="mx-auto max-w-6xl space-y-6">
          {/* Back link */}
          <Button asChild variant="ghost" size="sm" className="-ml-2">
            <Link href="/admin/sales-reps">
              <ArrowLeft className="mr-1 h-4 w-4" />
              Back to Team
            </Link>
          </Button>

          {/* Profile Card */}
          <div className="rounded-lg border p-6">
            <div className="flex items-start justify-between">
              <div className="space-y-1">
                <div className="flex items-center gap-3">
                  <h2 className="text-2xl font-bold">{rep.name}</h2>
                  <Badge variant={ROLE_COLORS[rep.employee_role]}>
                    {ROLE_LABELS[rep.employee_role] || rep.employee_role}
                  </Badge>
                  <Badge variant={rep.is_active ? "default" : "destructive"}>
                    {rep.is_active ? "Active" : "Inactive"}
                  </Badge>
                </div>
                <div className="flex items-center gap-4 text-sm text-muted-foreground mt-2">
                  <span className="flex items-center gap-1">
                    <Mail className="h-3.5 w-3.5" />
                    {rep.email}
                  </span>
                  {rep.phone && (
                    <span className="flex items-center gap-1">
                      <Phone className="h-3.5 w-3.5" />
                      {rep.phone}
                    </span>
                  )}
                  <span className="flex items-center gap-1">
                    <MapPin className="h-3.5 w-3.5" />
                    {rep.office}
                    {rep.territory ? ` — ${rep.territory}` : ""}
                  </span>
                </div>
              </div>
              <div className="text-right text-sm text-muted-foreground">
                <p>Commission: <span className="font-medium text-foreground">{rep.commission_rate}%</span></p>
                <p>Member since {formatDate(rep.created_at)}</p>
              </div>
            </div>
          </div>

          {/* KPI Cards */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <StatCard
              title="Total Quotes"
              value={String(kpis.totalQuotes)}
              icon={<FileText className="h-5 w-5 text-indigo-500" />}
            />
            <StatCard
              title="Total Revenue"
              value={formatCurrency(kpis.totalRevenue)}
              icon={<DollarSign className="h-5 w-5 text-green-500" />}
            />
            <StatCard
              title="Customers"
              value={String(kpis.totalCustomers)}
              icon={<Users className="h-5 w-5 text-blue-500" />}
            />
            <StatCard
              title="Conversion Rate"
              value={`${kpis.conversionRate.toFixed(1)}%`}
              subtitle={`${kpis.acceptedQuotes} accepted`}
              icon={<TrendingUp className="h-5 w-5 text-amber-500" />}
            />
          </div>

          {/* Quote Status Breakdown */}
          {Object.keys(kpis.statusCounts).length > 0 && (
            <div className="rounded-lg border p-4">
              <h3 className="text-sm font-semibold mb-3">Quote Status Breakdown</h3>
              <div className="flex flex-wrap gap-3">
                {Object.entries(kpis.statusCounts).map(([status, count]) => (
                  <div key={status} className="flex items-center gap-2">
                    <Badge variant={STATUS_COLORS[status]}>{status}</Badge>
                    <span className="text-sm font-medium">{count}</span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Monthly Revenue Chart */}
          <div className="rounded-lg border p-4">
            <h3 className="text-sm font-semibold mb-4">Monthly Revenue (Last 6 Months)</h3>
            <div className="flex items-end gap-2 h-40">
              {monthlyRevenue.map((m) => (
                <div
                  key={m.month}
                  className="flex-1 flex flex-col items-center gap-1"
                >
                  <span className="text-xs text-muted-foreground">
                    {m.count > 0 ? formatCurrency(m.revenue) : ""}
                  </span>
                  <div
                    className="w-full bg-primary/20 rounded-t relative"
                    style={{
                      height: `${Math.max((m.revenue / maxRevenue) * 100, 2)}%`,
                    }}
                  >
                    <div
                      className="absolute inset-0 bg-primary rounded-t"
                      style={{ opacity: m.revenue > 0 ? 1 : 0.2 }}
                    />
                  </div>
                  <span className="text-xs text-muted-foreground">
                    {m.month}
                  </span>
                </div>
              ))}
            </div>
          </div>

          {/* Quotes and Customers side by side on large screens */}
          <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
            {/* Quotes */}
            <div className="rounded-lg border">
              <div className="p-4 border-b flex items-center justify-between">
                <h3 className="text-sm font-semibold">Quotes ({filteredQuotes.length})</h3>
                <Select value={quoteStatusFilter} onValueChange={setQuoteStatusFilter}>
                  <SelectTrigger className="w-32 h-8 text-xs">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="all">All</SelectItem>
                    <SelectItem value="draft">Draft</SelectItem>
                    <SelectItem value="sent">Sent</SelectItem>
                    <SelectItem value="accepted">Accepted</SelectItem>
                    <SelectItem value="expired">Expired</SelectItem>
                  </SelectContent>
                </Select>
              </div>
              {filteredQuotes.length === 0 ? (
                <div className="p-6 text-center text-sm text-muted-foreground">
                  No quotes found.
                </div>
              ) : (
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Quote #</TableHead>
                      <TableHead>Customer</TableHead>
                      <TableHead>Status</TableHead>
                      <TableHead className="text-right">Total</TableHead>
                      <TableHead>Date</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {filteredQuotes.map((q) => (
                      <TableRow key={q.id}>
                        <TableCell>
                          <Link
                            href={`/quotes/${q.id}`}
                            className="font-medium text-primary hover:underline"
                          >
                            {q.quote_number}
                          </Link>
                        </TableCell>
                        <TableCell>{q.customer_name || "---"}</TableCell>
                        <TableCell>
                          <Badge variant={STATUS_COLORS[q.status]}>
                            {q.status}
                          </Badge>
                        </TableCell>
                        <TableCell className="text-right font-medium">
                          {formatCurrency(q.total)}
                        </TableCell>
                        <TableCell className="text-muted-foreground text-xs">
                          {formatDate(q.created_at)}
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              )}
            </div>

            {/* Customers */}
            <div className="rounded-lg border">
              <div className="p-4 border-b">
                <h3 className="text-sm font-semibold">Assigned Customers ({customers.length})</h3>
              </div>
              {customers.length === 0 ? (
                <div className="p-6 text-center text-sm text-muted-foreground">
                  No customers assigned.
                </div>
              ) : (
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Name</TableHead>
                      <TableHead>Email</TableHead>
                      <TableHead>Location</TableHead>
                      <TableHead>Added</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {customers.map((c) => (
                      <TableRow key={c.id}>
                        <TableCell>
                          <Link
                            href={`/customers/${c.id}`}
                            className="font-medium text-primary hover:underline"
                          >
                            {c.name}
                          </Link>
                        </TableCell>
                        <TableCell>{c.email || "---"}</TableCell>
                        <TableCell>
                          {[c.city, c.state].filter(Boolean).join(", ") || "---"}
                        </TableCell>
                        <TableCell className="text-muted-foreground text-xs">
                          {formatDate(c.created_at)}
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              )}
            </div>
          </div>

          {/* Activity Timeline */}
          <div className="rounded-lg border">
            <div className="p-4 border-b flex items-center gap-2">
              <Activity className="h-4 w-4 text-muted-foreground" />
              <h3 className="text-sm font-semibold">Recent Activity</h3>
            </div>
            {activity.length === 0 ? (
              <div className="p-6 text-center text-sm text-muted-foreground">
                No activity recorded yet.
              </div>
            ) : (
              <div className="divide-y">
                {activity.map((a) => (
                  <div key={a.id} className="px-4 py-3 flex items-start justify-between">
                    <div>
                      <p className="text-sm font-medium">
                        {ACTION_LABELS[a.action] || a.action}
                      </p>
                      {a.resource_type && (
                        <p className="text-xs text-muted-foreground">
                          {a.resource_type}
                          {a.resource_id ? ` #${a.resource_id.slice(0, 8)}` : ""}
                        </p>
                      )}
                    </div>
                    <span className="text-xs text-muted-foreground whitespace-nowrap">
                      {formatDate(a.created_at)}
                    </span>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>
    </>
  );
}

function StatCard({
  title,
  value,
  subtitle,
  icon,
}: {
  title: string;
  value: string;
  subtitle?: string;
  icon: React.ReactNode;
}) {
  return (
    <div className="rounded-lg border p-4 space-y-2">
      <div className="flex items-center justify-between">
        <span className="text-sm text-muted-foreground">{title}</span>
        {icon}
      </div>
      <p className="text-2xl font-bold">{value}</p>
      {subtitle && (
        <p className="text-xs text-muted-foreground">{subtitle}</p>
      )}
    </div>
  );
}
