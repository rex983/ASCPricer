"use client";

import { useCallback, useEffect, useState } from "react";
import Link from "next/link";
import { AppHeader } from "@/components/layout/app-header";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
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
import { formatCurrency, formatDate, STATUS_COLORS } from "@/lib/utils";
import {
  Calendar,
  DollarSign,
  FileText,
  Loader2,
  TrendingUp,
  Users,
} from "lucide-react";

interface DashboardData {
  totalCustomers: number;
  totalQuotes: number;
  totalRevenue: number;
  statusCounts: Record<string, number>;
  recentQuotes: {
    id: string;
    quote_number: string;
    status: string;
    total: number;
    created_at: string;
  }[];
  recentCustomers: {
    id: string;
    name: string;
    created_at: string;
  }[];
  monthlyRevenue: {
    month: string;
    revenue: number;
    count: number;
  }[];
}

const RECENT_LIMITS = [10, 25, 50, 100] as const;
type RecentLimit = (typeof RECENT_LIMITS)[number];

function toLocalDate(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function getPresetRange(preset: string): { start: string; end: string } {
  const now = new Date();
  const today = toLocalDate(now);
  switch (preset) {
    case "this_week": {
      const day = now.getDay();
      const sun = new Date(now);
      sun.setDate(now.getDate() - day);
      return { start: toLocalDate(sun), end: today };
    }
    case "last_week": {
      const day = now.getDay();
      const prevSun = new Date(now);
      prevSun.setDate(now.getDate() - day - 7);
      const prevSat = new Date(prevSun);
      prevSat.setDate(prevSun.getDate() + 6);
      return { start: toLocalDate(prevSun), end: toLocalDate(prevSat) };
    }
    case "this_month":
      return {
        start: `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-01`,
        end: today,
      };
    case "last_month": {
      const first = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      const last = new Date(now.getFullYear(), now.getMonth(), 0);
      return { start: toLocalDate(first), end: toLocalDate(last) };
    }
    case "this_quarter": {
      const qStart = new Date(now.getFullYear(), Math.floor(now.getMonth() / 3) * 3, 1);
      return { start: toLocalDate(qStart), end: today };
    }
    case "this_year":
      return { start: `${now.getFullYear()}-01-01`, end: today };
    case "last_year":
      return {
        start: `${now.getFullYear() - 1}-01-01`,
        end: `${now.getFullYear() - 1}-12-31`,
      };
    case "all_time":
      return { start: "2020-01-01", end: today };
    default:
      return {
        start: `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-01`,
        end: today,
      };
  }
}

const PRESETS = [
  { value: "this_week", label: "This Week" },
  { value: "last_week", label: "Last Week" },
  { value: "this_month", label: "This Month" },
  { value: "last_month", label: "Last Month" },
  { value: "this_quarter", label: "This Quarter" },
  { value: "this_year", label: "This Year" },
  { value: "last_year", label: "Last Year" },
  { value: "all_time", label: "All Time" },
  { value: "custom", label: "Custom" },
] as const;

export default function DashboardPage() {
  const defaultRange = getPresetRange("this_month");
  const [data, setData] = useState<DashboardData | null>(null);
  const [loading, setLoading] = useState(true);
  const [recentLimit, setRecentLimit] = useState<RecentLimit>(10);
  const [preset, setPreset] = useState("this_month");
  const [startDate, setStartDate] = useState(defaultRange.start);
  const [endDate, setEndDate] = useState(defaultRange.end);

  const fetchDashboard = useCallback(async (limit: RecentLimit, start: string, end: string) => {
    setLoading(true);
    try {
      const res = await fetch(
        `/api/dashboard?limit=${limit}&startDate=${start}&endDate=${end}`
      );
      const json = await res.json();
      if (json.totalQuotes !== undefined) setData(json);
    } catch {
      /* ignore */
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchDashboard(recentLimit, startDate, endDate);
  }, [fetchDashboard, recentLimit, startDate, endDate]);

  const handlePresetChange = (value: string) => {
    setPreset(value);
    if (value !== "custom") {
      const range = getPresetRange(value);
      setStartDate(range.start);
      setEndDate(range.end);
    }
  };

  if (loading) {
    return (
      <>
        <AppHeader title="Dashboard" />
        <div className="flex-1 flex items-center justify-center py-12">
          <Loader2 className="h-8 w-8 animate-spin text-muted-foreground" />
        </div>
      </>
    );
  }

  if (!data) {
    return (
      <>
        <AppHeader title="Dashboard" />
        <div className="flex-1 p-6 text-center text-muted-foreground">
          Failed to load dashboard data.
        </div>
      </>
    );
  }

  // Find max revenue for simple bar chart scaling
  const maxRevenue = Math.max(...data.monthlyRevenue.map((m) => m.revenue), 1);

  const presetLabel = PRESETS.find((p) => p.value === preset)?.label ?? "Custom";

  return (
    <>
      <AppHeader title="Dashboard" />
      <div className="flex-1 p-4 sm:p-6">
        <div className="mx-auto max-w-5xl space-y-4 sm:space-y-6">
          {/* Date Range Picker */}
          <div className="flex flex-wrap items-center gap-2 sm:gap-3 rounded-lg border p-3">
            <Calendar className="hidden sm:block h-4 w-4 text-muted-foreground shrink-0" />
            <Select value={preset} onValueChange={handlePresetChange}>
              <SelectTrigger className="h-9 w-full sm:w-36 text-sm">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                {PRESETS.map((p) => (
                  <SelectItem key={p.value} value={p.value}>
                    {p.label}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
            <div className="flex w-full sm:w-auto items-center gap-2">
              <Input
                type="date"
                className="h-9 flex-1 sm:w-[150px] text-sm"
                value={startDate}
                onChange={(e) => {
                  setStartDate(e.target.value);
                  setPreset("custom");
                }}
              />
              <span className="text-sm text-muted-foreground">to</span>
              <Input
                type="date"
                className="h-9 flex-1 sm:w-[150px] text-sm"
                value={endDate}
                onChange={(e) => {
                  setEndDate(e.target.value);
                  setPreset("custom");
                }}
              />
            </div>
            {preset !== "this_month" && (
              <Button
                variant="ghost"
                size="sm"
                className="text-xs"
                onClick={() => handlePresetChange("this_month")}
              >
                Reset
              </Button>
            )}
          </div>

          {/* Stat Cards */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <StatCard
              title="Customers"
              value={String(data.totalCustomers)}
              icon={<Users className="h-5 w-5 text-blue-500" />}
            />
            <StatCard
              title="Quotes"
              value={String(data.totalQuotes)}
              icon={<FileText className="h-5 w-5 text-indigo-500" />}
            />
            <StatCard
              title="Total Revenue"
              value={formatCurrency(data.totalRevenue)}
              icon={<DollarSign className="h-5 w-5 text-green-500" />}
            />
            <StatCard
              title="Avg per Quote"
              value={
                data.totalQuotes > 0
                  ? formatCurrency(data.totalRevenue / data.totalQuotes)
                  : "$0.00"
              }
              icon={<TrendingUp className="h-5 w-5 text-amber-500" />}
            />
          </div>

          {/* Quote Status Breakdown */}
          {Object.keys(data.statusCounts).length > 0 && (
            <div className="rounded-lg border p-4">
              <h3 className="text-sm font-semibold mb-3">Quote Status Breakdown</h3>
              <div className="flex flex-wrap gap-3">
                {Object.entries(data.statusCounts).map(([status, count]) => (
                  <div key={status} className="flex items-center gap-2">
                    <Badge variant={STATUS_COLORS[status]}>{status}</Badge>
                    <span className="text-sm font-medium">{count}</span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Monthly Revenue Chart (simple bar chart) */}
          <div className="rounded-lg border p-4">
            <h3 className="text-sm font-semibold mb-4">
              Monthly Revenue{data.monthlyRevenue.length <= 1 ? "" : ` (${data.monthlyRevenue[0]?.month} – ${data.monthlyRevenue[data.monthlyRevenue.length - 1]?.month})`}
            </h3>
            <div className="flex items-end gap-2 h-40">
              {data.monthlyRevenue.map((m) => (
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
                      style={{
                        opacity: m.revenue > 0 ? 1 : 0.2,
                      }}
                    />
                  </div>
                  <span className="text-xs text-muted-foreground">
                    {m.month}
                  </span>
                </div>
              ))}
            </div>
          </div>

          {/* Recent Activity */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* Recent Quotes */}
            <div className="rounded-lg border">
              <div className="p-4 border-b flex items-center justify-between gap-2">
                <h3 className="text-sm font-semibold">Recent Quotes</h3>
                <Select
                  value={String(recentLimit)}
                  onValueChange={(v) => setRecentLimit(Number(v) as RecentLimit)}
                >
                  <SelectTrigger className="h-8 w-28 text-xs">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    {RECENT_LIMITS.map((n) => (
                      <SelectItem key={n} value={String(n)}>
                        Last {n}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              {data.recentQuotes.length === 0 ? (
                <div className="p-6 text-center text-sm text-muted-foreground">
                  No quotes yet.
                </div>
              ) : (
                <div className="overflow-x-auto">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Quote #</TableHead>
                      <TableHead>Status</TableHead>
                      <TableHead className="text-right">Total</TableHead>
                      <TableHead className="hidden sm:table-cell">Date</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {data.recentQuotes.map((q) => (
                      <TableRow key={q.id}>
                        <TableCell>
                          <Link
                            href={`/quotes/${q.id}`}
                            className="font-medium text-primary hover:underline"
                          >
                            {q.quote_number}
                          </Link>
                        </TableCell>
                        <TableCell>
                          <Badge variant={STATUS_COLORS[q.status]}>
                            {q.status}
                          </Badge>
                        </TableCell>
                        <TableCell className="text-right">
                          {formatCurrency(q.total)}
                        </TableCell>
                        <TableCell className="hidden sm:table-cell text-muted-foreground text-xs">
                          {formatDate(q.created_at)}
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
                </div>
              )}
            </div>

            {/* Recent Customers */}
            <div className="rounded-lg border">
              <div className="p-4 border-b">
                <h3 className="text-sm font-semibold">Recent Customers</h3>
              </div>
              {data.recentCustomers.length === 0 ? (
                <div className="p-6 text-center text-sm text-muted-foreground">
                  No customers yet.
                </div>
              ) : (
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Name</TableHead>
                      <TableHead>Added</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {data.recentCustomers.map((c) => (
                      <TableRow key={c.id}>
                        <TableCell>
                          <Link
                            href={`/customers/${c.id}`}
                            className="font-medium text-primary hover:underline"
                          >
                            {c.name}
                          </Link>
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
        </div>
      </div>
    </>
  );
}

function StatCard({
  title,
  value,
  icon,
}: {
  title: string;
  value: string;
  icon: React.ReactNode;
}) {
  return (
    <div className="rounded-lg border p-4 space-y-2">
      <div className="flex items-center justify-between">
        <span className="text-sm text-muted-foreground">{title}</span>
        {icon}
      </div>
      <p className="text-2xl font-bold">{value}</p>
    </div>
  );
}
