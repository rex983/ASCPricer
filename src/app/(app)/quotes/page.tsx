"use client";

import { useCallback, useEffect, useState } from "react";
import Link from "next/link";
import { useSession } from "next-auth/react";
import { AppHeader } from "@/components/layout/app-header";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { formatCurrency, formatDate } from "@/lib/utils";
import { FileText, Search, Trash2 } from "lucide-react";

interface QuoteSummary {
  id: string;
  quote_number: string;
  customer_name: string | null;
  customer_state: string | null;
  subtotal: number;
  total: number;
  office: string | null;
  created_at: string;
}

export default function QuotesPage() {
  const { data: session } = useSession();
  const [quotes, setQuotes] = useState<QuoteSummary[]>([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");

  const role = session?.user?.role;
  const showOffice = role === "admin" || role === "manager";
  const canDelete = role === "admin" || role === "manager";

  const fetchQuotes = useCallback(() => {
    setLoading(true);
    const params = new URLSearchParams();
    if (search) params.set("search", search);

    fetch(`/api/quotes?${params}`)
      .then((r) => r.json())
      .then((data) => {
        if (Array.isArray(data)) setQuotes(data);
      })
      .catch(() => {})
      .finally(() => setLoading(false));
  }, [search]);

  useEffect(() => {
    fetchQuotes();
  }, [fetchQuotes]);

  async function handleDelete(id: string, quoteNumber: string) {
    if (!confirm(`Delete quote ${quoteNumber}? This cannot be undone.`)) return;
    try {
      const res = await fetch(`/api/quotes/${id}`, { method: "DELETE" });
      if (res.ok) {
        setQuotes((prev) => prev.filter((q) => q.id !== id));
      } else {
        const data = await res.json();
        alert(data.error || "Failed to delete");
      }
    } catch {
      alert("Network error");
    }
  }

  return (
    <>
      <AppHeader title="Quotes" />
      <div className="flex-1 p-4 sm:p-6">
        <div className="mx-auto max-w-5xl space-y-4">
          {/* Filters */}
          <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-3">
            <div className="relative flex-1 max-w-sm">
              <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
              <Input
                placeholder="Search quotes..."
                className="pl-9"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
            </div>
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
            <div className="rounded-lg border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Quote #</TableHead>
                    <TableHead>Customer</TableHead>
                    <TableHead className="hidden sm:table-cell">State</TableHead>
                    {showOffice && <TableHead className="hidden md:table-cell">Office</TableHead>}
                    <TableHead className="text-right">Total</TableHead>
                    <TableHead className="hidden sm:table-cell">Date</TableHead>
                    {canDelete && <TableHead className="w-10" />}
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
                      <TableCell>{q.customer_name || "---"}</TableCell>
                      <TableCell className="hidden sm:table-cell">{q.customer_state || "---"}</TableCell>
                      {showOffice && (
                        <TableCell className="hidden md:table-cell">
                          {q.office ? (
                            <Badge variant="outline">{q.office}</Badge>
                          ) : (
                            "---"
                          )}
                        </TableCell>
                      )}
                      <TableCell className="text-right font-medium">
                        {formatCurrency(q.total)}
                      </TableCell>
                      <TableCell className="hidden sm:table-cell text-muted-foreground">
                        {formatDate(q.created_at)}
                      </TableCell>
                      {canDelete && (
                        <TableCell>
                          <Button
                            variant="ghost"
                            size="icon"
                            className="h-8 w-8 text-muted-foreground hover:text-destructive"
                            onClick={() => handleDelete(q.id, q.quote_number)}
                          >
                            <Trash2 className="h-4 w-4" />
                          </Button>
                        </TableCell>
                      )}
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
