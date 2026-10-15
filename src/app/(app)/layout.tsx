import { SidebarProvider, SidebarInset } from "@/components/ui/sidebar";
import { AppSidebar } from "@/components/layout/app-sidebar";
import { ImpersonationBanner } from "@/components/layout/impersonation-banner";
import { SessionProvider } from "@/components/providers/session-provider";
import { Tracker } from "@/components/analytics-tracker";

export default function AppLayout({ children }: { children: React.ReactNode }) {
  return (
    <SessionProvider>
      <Tracker />
      <SidebarProvider>
        <AppSidebar />
        <SidebarInset>
          <ImpersonationBanner />
          {children}
        </SidebarInset>
      </SidebarProvider>
    </SessionProvider>
  );
}
