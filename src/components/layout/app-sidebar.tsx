"use client";

import { useEffect, useState } from "react";
import Link from "next/link";
import { usePathname } from "next/navigation";
import { useSession } from "next-auth/react";
import { Calculator, FileText, Upload, Map, History, SlidersHorizontal, Users, LayoutDashboard, ShieldCheck } from "lucide-react";
import {
  Sidebar,
  SidebarContent,
  SidebarGroup,
  SidebarGroupLabel,
  SidebarGroupContent,
  SidebarMenu,
  SidebarMenuItem,
  SidebarMenuButton,
  SidebarHeader,
  SidebarFooter,
} from "@/components/ui/sidebar";
import { UserNav } from "@/components/layout/user-nav";
import { Logo } from "@/components/layout/logo";

const navItems = [
  { title: "Dashboard", href: "/dashboard", icon: LayoutDashboard },
  { title: "Calculator", href: "/calculator", icon: Calculator },
  { title: "Quotes", href: "/quotes", icon: FileText },
  { title: "Customers", href: "/customers", icon: Users },
];

const adminItems = [
  { title: "Users", href: "/admin/users", icon: ShieldCheck },
  { title: "Upload Pricing", href: "/admin/upload", icon: Upload },
  { title: "Regions", href: "/admin/regions", icon: Map },
  { title: "Variables", href: "/admin/variables", icon: SlidersHorizontal },
  { title: "Audit Log", href: "/admin/audit-log", icon: History },
];

export function AppSidebar() {
  const pathname = usePathname();
  const { data: session } = useSession();
  const role = session?.user?.role;
  const isAdminOrManager = role === "admin" || role === "manager";
  const [impersonating, setImpersonating] = useState(false);

  useEffect(() => {
    if (!isAdminOrManager) return;
    fetch("/api/view-as")
      .then((r) => r.json())
      .then((d) => setImpersonating(!!d.isImpersonating))
      .catch(() => {});
  }, [isAdminOrManager, pathname]);

  const showAdmin = isAdminOrManager && !impersonating;

  return (
    <Sidebar>
      <SidebarHeader className="border-b px-4 py-3">
        <Link href="/calculator">
          <Logo size="small" />
        </Link>
      </SidebarHeader>
      <SidebarContent>
        <SidebarGroup>
          <SidebarGroupLabel>Main</SidebarGroupLabel>
          <SidebarGroupContent>
            <SidebarMenu>
              {navItems.map((item) => (
                <SidebarMenuItem key={item.href}>
                  <SidebarMenuButton
                    asChild
                    isActive={pathname.startsWith(item.href)}
                  >
                    <Link href={item.href}>
                      <item.icon className="h-4 w-4" />
                      <span>{item.title}</span>
                    </Link>
                  </SidebarMenuButton>
                </SidebarMenuItem>
              ))}
            </SidebarMenu>
          </SidebarGroupContent>
        </SidebarGroup>

        {showAdmin && (
          <SidebarGroup>
            <SidebarGroupLabel>Admin</SidebarGroupLabel>
            <SidebarGroupContent>
              <SidebarMenu>
                {adminItems.map((item) => (
                  <SidebarMenuItem key={item.href}>
                    <SidebarMenuButton
                      asChild
                      isActive={pathname.startsWith(item.href)}
                    >
                      <Link href={item.href}>
                        <item.icon className="h-4 w-4" />
                        <span>{item.title}</span>
                      </Link>
                    </SidebarMenuButton>
                  </SidebarMenuItem>
                ))}
              </SidebarMenu>
            </SidebarGroupContent>
          </SidebarGroup>
        )}
      </SidebarContent>
      <SidebarFooter className="border-t p-4">
        <UserNav />
      </SidebarFooter>
    </Sidebar>
  );
}
