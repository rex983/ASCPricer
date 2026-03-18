"use client";

import { useEffect, useState } from "react";
import { useParams, useRouter } from "next/navigation";
import Link from "next/link";
import { useSession } from "next-auth/react";
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
import { formatCurrency, formatDate } from "@/lib/utils";
import { ArrowLeft, FileText, Trash2 } from "lucide-react";

interface CustomerDetail {
  id: string;
  name: string;
  email: string | null;
  phone: string | null;
  address: string | null;
  city: string | null;
  state: string | null;
  zip: string | null;
  office: string;
  notes: string | null;
  created_at: string;
  quotes: {
    id: string;
    quote_number: string;
    status: string;
    subtotal: number;
    total: number;
    created_at: string;
  }[];
}

const STATUS_COLORS: Record<string, string> = {
  draft: "secondary",
  sent: "default",
  accepted: "default",
  expired: "destructive",
};

export default function CustomerDetailPage() {
  const { id } = useParams<{ id: string }>();
  const router = useRouter();
  const { data: session } = useSession();
  const [customer, setCustomer] = useState<CustomerDetail | null>(null);
  const [loading, setLoading] = useState(true);

  const role = session?.user?.role;
  const canDelete = role === "admin" || role === "manager" || role === "sales_rep";

  useEffect(() => {
    fetch(`/api/customers/${id}`)
      .then((r) => r.json())
      .then((data) => {
        if (data.id) setCustomer(data);
      })
      .catch(() => {})
      .finally(() => setLoading(false));
  }, [id]);

  async function handleDelete() {
    if (!customer) return;
    if (!confirm(`Delete customer "${customer.name}"? This cannot be undone.`)) return;
    try {
      const res = await fetch(`/api/customers/${id}`, { method: "DELETE" });
      if (res.ok) {
        router.push("/customers");
      } else {
        const data = await res.json();
        alert(data.error || "Failed to delete");
      }
    } catch {
      alert("Network error");
    }
  }

  if (loading) {
    return (
      <>
        <AppHeader title="Customer" />
        <div className="flex-1 p-6 text-center text-muted-foreground">Loading...</div>
      </>
    );
  }

  if (!customer) {
    return (
      <>
        <AppHeader title="Customer" />
        <div className="flex-1 p-6 text-center text-muted-foreground">Customer not found.</div>
      </>
    );
  }

  const location = [customer.address, customer.city, customer.state, customer.zip]
    .filter(Boolean)
    .join(", ");

  return (
    <>
      <AppHeader title={customer.name} />
      <div className="flex-1 p-6">
        <div className="mx-auto max-w-4xl space-y-6">
          {/* Back + Actions */}
          <div className="flex items-center justify-between">
            <Button variant="ghost" size="sm" asChild>
              <Link href="/customers">
                <ArrowLeft className="mr-1 h-4 w-4" /> All Customers
              </Link>
            </Button>
            {canDelete && (
              <Button variant="destructive" size="sm" onClick={handleDelete}>
                <Trash2 className="mr-1 h-4 w-4" /> Delete Customer
              </Button>
            )}
          </div>

          {/* Customer Info */}
          <div className="rounded-lg border p-5 grid grid-cols-2 md:grid-cols-3 gap-4 text-sm">
            <div>
              <p className="text-muted-foreground">Name</p>
              <p className="font-medium">{customer.name}</p>
            </div>
            <div>
              <p className="text-muted-foreground">Email</p>
              <p className="font-medium">{customer.email || "---"}</p>
            </div>
            <div>
              <p className="text-muted-foreground">Phone</p>
              <p className="font-medium">{customer.phone || "---"}</p>
            </div>
            <div className="col-span-2">
              <p className="text-muted-foreground">Address</p>
              <p className="font-medium">{location || "---"}</p>
            </div>
            <div>
              <p className="text-muted-foreground">Office</p>
              <Badge variant="outline">{customer.office}</Badge>
            </div>
            {customer.notes && (
              <div className="col-span-full">
                <p className="text-muted-foreground">Notes</p>
                <p className="font-medium whitespace-pre-wrap">{customer.notes}</p>
              </div>
            )}
          </div>

          {/* Quotes */}
          <div className="space-y-2">
            <h2 className="text-lg font-semibold">
              Quotes ({customer.quotes.length})
            </h2>
            {customer.quotes.length === 0 ? (
              <div className="rounded-lg border border-dashed p-8 text-center text-muted-foreground">
                <FileText className="mx-auto h-8 w-8 mb-2" />
                <p>No quotes for this customer yet.</p>
              </div>
            ) : (
              <div className="rounded-lg border">
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
                    {customer.quotes.map((q) => (
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
                          <Badge
                            variant={
                              STATUS_COLORS[q.status] as
                                | "default"
                                | "secondary"
                                | "destructive"
                            }
                          >
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
      </div>
    </>
  );
}
