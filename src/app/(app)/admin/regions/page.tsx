import { AppHeader } from "@/components/layout/app-header";

export default function RegionsPage() {
  return (
    <>
      <AppHeader title="Regions" />
      <div className="flex-1 p-6">
        <div className="rounded-lg border border-dashed p-12 text-center text-muted-foreground">
          <h2 className="text-lg font-medium">Region Management</h2>
          <p className="mt-2">
            Add, edit, and manage pricing regions.
          </p>
          <p className="mt-1 text-sm">Coming in Phase 7</p>
        </div>
      </div>
    </>
  );
}
