"use client";

import { useEffect, useState } from "react";
import Link from "next/link";
import { AppHeader } from "@/components/layout/app-header";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { formatCurrency, formatDate } from "@/lib/utils";
import { FileText, Search } from "lucide-react";

interface QuoteSummary {
  id: string;
  quote_number: string;
  status: string;
  customer_name: string | null;
  customer_state: string | null;
  subtotal: number;
  total: number;
  created_at: string;
}

const STATUS_COLORS: Record<string, string> = {
  draft: "secondary",
  sent: "default",
  accepted: "default",
  expired: "destructive",
};

export default function QuotesPage() {
  const [quotes, setQuotes] = useState<QuoteSummary[]>([]);
  const [loading, setLoading] = useState(true);
  const [statusFilter, setStatusFilter] = useState("all");
  const [search, setSearch] = useState("");

  useEffect(() => {
    setLoading(true);
    const params = new URLSearchParams();
    if (statusFilter !== "all") params.set("status", statusFilter);
    if (search) params.set("search", search);

    fetch(`/api/quotes?${params}`)
      .then((r) => r.json())
      .then((data) => {
        if (Array.isArray(data)) setQuotes(data);
      })
      .catch(() => {})
      .finally(() => setLoading(false));
  }, [statusFilter, search]);

  return (
    <>
      <AppHeader title="Quotes" />
      <div className="flex-1 p-6">
        <div className="mx-auto max-w-5xl space-y-4">
          {/* Filters */}
          <div className="flex items-center gap-3">
            <div className="relative flex-1 max-w-sm">
              <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
              <Input
                placeholder="Search quotes..."
                className="pl-9"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
            </div>
            <Select value={statusFilter} onValueChange={setStatusFilter}>
              <SelectTrigger className="w-36">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">All Statuses</SelectItem>
                <SelectItem value="draft">Draft</SelectItem>
                <SelectItem value="sent">Sent</SelectItem>
                <SelectItem value="accepted">Accepted</SelectItem>
                <SelectItem value="expired">Expired</SelectItem>
              </SelectContent>
            </Select>
            <Button asChild>
              <Link href="/calculator">New Quote</Link>
            </Button>
          </div>

          {/* Table */}
          {loading ? (
            <div className="text-center text-muted-foreground py-12">Loading...</div>
          ) : quotes.length === 0 ? (
            <div className="rounded-lg border border-dashed p-12 text-center text-muted-foreground">
              <FileText className="mx-auto h-10 w-10 mb-2" />
              <p>No quotes yet. Create one from the calculator.</p>
            </div>
          ) : (
            <div className="rounded-lg border">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Quote #</TableHead>
                    <TableHead>Customer</TableHead>
                    <TableHead>State</TableHead>
                    <TableHead>Status</TableHead>
                    <TableHead className="text-right">Total</TableHead>
                    <TableHead>Date</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {quotes.map((q) => (
                    <TableRow key={q.id}>
                      <TableCell>
                        <Link
                          href={`/quotes/${q.id}`}
                          className="font-medium text-primary hover:underline"
                        >
                          {q.quote_number}
                        </Link>
                      </TableCell>
                      <TableCell>{q.customer_name || "—"}</TableCell>
                      <TableCell>{q.customer_state || "—"}</TableCell>
                      <TableCell>
                        <Badge variant={STATUS_COLORS[q.status] as "default" | "secondary" | "destructive"}>
                          {q.status}
                        </Badge>
                      </TableCell>
                      <TableCell className="text-right font-medium">
                        {formatCurrency(q.total)}
                      </TableCell>
                      <TableCell className="text-muted-foreground">
                        {formatDate(q.created_at)}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          )}
        </div>
      </div>
    </>
  );
}
