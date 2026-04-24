"use client";

import { useCallback, useEffect, useState } from "react";
import Link from "next/link";
import { useSession } from "next-auth/react";
import { AppHeader } from "@/components/layout/app-header";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from "@/components/ui/dialog";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { canDeleteRecord } from "@/lib/utils";
import { Plus, Search, Trash2, Users } from "lucide-react";

interface Customer {
  id: string;
  name: string;
  email: string | null;
  phone: string | null;
  city: string | null;
  state: string | null;
  office: string;
  created_at: string;
}

export default function CustomersPage() {
  const { data: session } = useSession();
  const [customers, setCustomers] = useState<Customer[]>([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [dialogOpen, setDialogOpen] = useState(false);

  const userOffice = session?.user?.office;
  const userRole = session?.user?.role;
  const canDelete = canDeleteRecord(userRole);

  const fetchCustomers = useCallback(async () => {
    setLoading(true);
    const params = new URLSearchParams();
    if (search) params.set("search", search);
    try {
      const r = await fetch(`/api/customers?${params}`);
      const data = await r.json();
      if (Array.isArray(data)) setCustomers(data);
    } catch {
      // ignore
    } finally {
      setLoading(false);
    }
  }, [search]);

  useEffect(() => {
    fetchCustomers();
  }, [fetchCustomers]);

  async function handleDelete(id: string, name: string) {
    if (!confirm(`Delete customer "${name}"? This cannot be undone.`)) return;
    try {
      const res = await fetch(`/api/customers/${id}`, { method: "DELETE" });
      if (res.ok) {
        setCustomers((prev) => prev.filter((x) => x.id !== id));
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
      <AppHeader title="Customers" />
      <div className="flex-1 p-4 sm:p-6">
        <div className="mx-auto max-w-5xl space-y-4">
          <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-3">
            <div className="relative flex-1 max-w-sm">
              <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
              <Input
                placeholder="Search customers..."
                className="pl-9"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
            </div>
            <Dialog open={dialogOpen} onOpenChange={setDialogOpen}>
              <DialogTrigger asChild>
                <Button>
                  <Plus className="mr-1 h-4 w-4" /> New Customer
                </Button>
              </DialogTrigger>
              <DialogContent>
                <DialogHeader>
                  <DialogTitle>New Customer</DialogTitle>
                </DialogHeader>
                <NewCustomerForm
                  defaultOffice={userOffice}
                  showOfficeSelect={userRole === "admin"}
                  onSaved={() => {
                    setDialogOpen(false);
                    fetchCustomers();
                  }}
                />
              </DialogContent>
            </Dialog>
          </div>

          {loading ? (
            <div className="text-center text-muted-foreground py-12">
              Loading...
            </div>
          ) : customers.length === 0 ? (
            <div className="rounded-lg border border-dashed p-12 text-center text-muted-foreground">
              <Users className="mx-auto h-10 w-10 mb-2" />
              <p>No customers yet.</p>
            </div>
          ) : (
            <div className="rounded-lg border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Name</TableHead>
                    <TableHead className="hidden sm:table-cell">Email</TableHead>
                    <TableHead className="hidden md:table-cell">Phone</TableHead>
                    <TableHead className="hidden sm:table-cell">Location</TableHead>
                    <TableHead>Office</TableHead>
                    {canDelete && <TableHead className="w-10" />}
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
                      <TableCell className="hidden sm:table-cell">{c.email || "---"}</TableCell>
                      <TableCell className="hidden md:table-cell">{c.phone || "---"}</TableCell>
                      <TableCell className="hidden sm:table-cell">
                        {[c.city, c.state].filter(Boolean).join(", ") || "---"}
                      </TableCell>
                      <TableCell>
                        <Badge variant="outline">{c.office}</Badge>
                      </TableCell>
                      {canDelete && (
                        <TableCell>
                          <Button
                            variant="ghost"
                            size="icon"
                            className="h-8 w-8 text-muted-foreground hover:text-destructive"
                            onClick={() => handleDelete(c.id, c.name)}
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

/* ---------- New Customer Form ---------- */

function NewCustomerForm({
  defaultOffice,
  showOfficeSelect,
  onSaved,
}: {
  defaultOffice?: string;
  showOfficeSelect: boolean;
  onSaved: () => void;
}) {
  const [saving, setSaving] = useState(false);
  const [form, setForm] = useState({
    name: "",
    email: "",
    phone: "",
    address: "",
    city: "",
    state: "",
    zip: "",
    office: defaultOffice || "Harbor",
  });

  const set = (field: string, value: string) =>
    setForm((prev) => ({ ...prev, [field]: value }));

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!form.name.trim()) return;
    setSaving(true);
    try {
      const res = await fetch("/api/customers", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(form),
      });
      if (res.ok) {
        onSaved();
      } else {
        const data = await res.json();
        alert(data.error || "Failed to save");
      }
    } catch {
      alert("Network error");
    } finally {
      setSaving(false);
    }
  }

  return (
    <form onSubmit={handleSubmit} className="space-y-3">
      <div className="grid grid-cols-2 gap-3">
        <div className="col-span-2 space-y-1">
          <Label>Name *</Label>
          <Input
            value={form.name}
            onChange={(e) => set("name", e.target.value)}
            required
          />
        </div>
        <div className="space-y-1">
          <Label>Email</Label>
          <Input
            type="email"
            value={form.email}
            onChange={(e) => set("email", e.target.value)}
          />
        </div>
        <div className="space-y-1">
          <Label>Phone</Label>
          <Input
            value={form.phone}
            onChange={(e) => set("phone", e.target.value)}
          />
        </div>
        <div className="col-span-2 space-y-1">
          <Label>Address</Label>
          <Input
            value={form.address}
            onChange={(e) => set("address", e.target.value)}
          />
        </div>
        <div className="space-y-1">
          <Label>City</Label>
          <Input
            value={form.city}
            onChange={(e) => set("city", e.target.value)}
          />
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div className="space-y-1">
            <Label>State</Label>
            <Input
              value={form.state}
              onChange={(e) => set("state", e.target.value)}
            />
          </div>
          <div className="space-y-1">
            <Label>ZIP</Label>
            <Input
              value={form.zip}
              onChange={(e) => set("zip", e.target.value)}
            />
          </div>
        </div>
        <div className="col-span-2 space-y-1">
          <Label>Office *</Label>
          {showOfficeSelect ? (
            <Select value={form.office} onValueChange={(v) => set("office", v)}>
              <SelectTrigger>
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="Harbor">Harbor</SelectItem>
                <SelectItem value="Marion">Marion</SelectItem>
              </SelectContent>
            </Select>
          ) : (
            <Input value={form.office} disabled />
          )}
        </div>
      </div>
      <Button type="submit" disabled={saving} className="w-full">
        {saving ? "Saving..." : "Create Customer"}
      </Button>
    </form>
  );
}
