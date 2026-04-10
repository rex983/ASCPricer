"use client";

import { useCallback, useEffect, useState } from "react";
import Link from "next/link";
import { AppHeader } from "@/components/layout/app-header";
import { Badge } from "@/components/ui/badge";
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

export default function DashboardPage() {
  const [data, setData] = useState<DashboardData | null>(null);
  const [loading, setLoading] = useState(true);

  const fetchDashboard = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch("/api/dashboard");
      const json = await res.json();
      if (json.totalQuotes !== undefined) setData(json);
    } catch {
      /* ignore */
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchDashboard();
  }, [fetchDashboard]);

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

  return (
    <>
      <AppHeader title="Dashboard" />
      <div className="flex-1 p-6">
        <div className="mx-auto max-w-5xl space-y-6">
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
            <h3 className="text-sm font-semibold mb-4">Monthly Revenue (Last 6 Months)</h3>
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
              <div className="p-4 border-b">
                <h3 className="text-sm font-semibold">Recent Quotes</h3>
              </div>
              {data.recentQuotes.length === 0 ? (
                <div className="p-6 text-center text-sm text-muted-foreground">
                  No quotes yet.
                </div>
              ) : (
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Quote #</TableHead>
                      <TableHead>Status</TableHead>
                      <TableHead className="text-right">Total</TableHead>
                      <TableHead>Date</TableHead>
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
                        <TableCell className="text-muted-foreground text-xs">
                          {formatDate(q.created_at)}
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
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
