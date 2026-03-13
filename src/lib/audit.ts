import { createAdminClient } from "@/lib/supabase/admin";

interface AuditEntry {
  userId?: string | null;
  userEmail?: string | null;
  action: string;
  resourceType?: string;
  resourceId?: string;
  details?: Record<string, unknown>;
}

/** Insert an audit log entry. Fire-and-forget — never throws. */
export async function logAudit(entry: AuditEntry) {
  try {
    const supabase = createAdminClient();
    await supabase.from("asc_audit_log").insert({
      user_id: entry.userId ?? null,
      user_email: entry.userEmail ?? null,
      action: entry.action,
      resource_type: entry.resourceType ?? null,
      resource_id: entry.resourceId ?? null,
      details: entry.details ?? {},
    });
  } catch {
    // Never block the main flow for audit logging
  }
}
