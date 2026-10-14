import type { SupabaseClient } from "@supabase/supabase-js";
import { createAdminClient } from "@/lib/supabase/admin";

/**
 * Org vocabulary (roles / offices / departments) is owned by the BBD Launcher
 * (launcher_roles / launcher_offices / launcher_departments). Admins can add
 * new names there at any time, so nothing in this app may hard-code them.
 * Lists are read server-side and cached for up to 60s. If the launcher tables
 * aren't deployed yet, the previous built-in lists are used so deploy order
 * doesn't matter.
 */

export interface OrgOption {
  name: string;
  displayName: string;
}

export interface RoleOption extends OrgOption {
  isAdmin: boolean;
  actsAs: string | null;
}

const TTL_MS = 60_000;

// Fallbacks — only used while the launcher tables don't exist / can't be read.
const FALLBACK_OFFICES: OrgOption[] = [
  { name: "Sales", displayName: "Sales" },
  { name: "BST", displayName: "BST" },
  { name: "RnD", displayName: "RnD" },
];
const FALLBACK_DEPARTMENTS: OrgOption[] = [
  { name: "SALES TEAM", displayName: "SALES TEAM" },
  { name: "BST", displayName: "BST" },
  { name: "RnD", displayName: "RnD" },
];
const FALLBACK_ROLES: RoleOption[] = [
  ["admin", "Admin"],
  ["senior_manager", "Senior Manager"],
  ["junior_manager", "Junior Manager"],
  ["team_lead", "Team Lead"],
  ["sales_rep", "Sales Rep"],
  ["bst", "BST"],
  ["rnd", "R&D"],
  ["cancellations_dept", "Cancellations Dept"],
  ["revisions_dept", "Revisions Dept"],
].map(([name, displayName]) => ({
  name,
  displayName,
  isAdmin: name === "admin",
  actsAs: null,
}));

type Entry<T> = { at: number; value: T };
const cache = new Map<string, Entry<unknown>>();

async function cached<T>(key: string, load: () => Promise<T | null>, fallback: T): Promise<T> {
  const hit = cache.get(key) as Entry<T> | undefined;
  if (hit && Date.now() - hit.at < TTL_MS) return hit.value;
  let value: T | null = null;
  try {
    value = await load();
  } catch (err) {
    console.warn(`[launcher-org] ${key} read failed:`, err instanceof Error ? err.message : err);
  }
  if (value === null) {
    // Don't pin the fallback for long — retry soon once the tables exist.
    cache.set(key, { at: Date.now() - TTL_MS + 10_000, value: fallback });
    return fallback;
  }
  cache.set(key, { at: Date.now(), value });
  return value;
}

async function readOrgTable(table: string): Promise<OrgOption[] | null> {
  const sb = createAdminClient();
  const { data, error } = await sb
    .from(table)
    .select("name, display_name, sort_order")
    .order("sort_order", { ascending: true })
    .order("name", { ascending: true });
  if (error) {
    if (error.code !== "PGRST205" && error.code !== "42P01") {
      console.warn(`[launcher-org] ${table}: ${error.message}`);
    }
    return null;
  }
  return (data ?? []).map((r) => ({
    name: String(r.name),
    displayName: (r.display_name as string | null) || String(r.name),
  }));
}

export function getLauncherOffices(): Promise<OrgOption[]> {
  return cached("offices", () => readOrgTable("launcher_offices"), FALLBACK_OFFICES);
}

export function getLauncherDepartments(): Promise<OrgOption[]> {
  return cached("departments", () => readOrgTable("launcher_departments"), FALLBACK_DEPARTMENTS);
}

export function getLauncherRoles(): Promise<RoleOption[]> {
  return cached(
    "roles",
    async () => {
      const sb = createAdminClient();
      const { data, error } = await sb
        .from("launcher_roles")
        .select("name, display_name, is_admin, acts_as")
        .order("name", { ascending: true });
      if (error) {
        if (error.code !== "PGRST205" && error.code !== "42P01") {
          console.warn(`[launcher-org] launcher_roles: ${error.message}`);
        }
        return null;
      }
      return (data ?? []).map((r) => ({
        name: String(r.name),
        displayName: (r.display_name as string | null) || String(r.name),
        isAdmin: r.is_admin === true,
        actsAs: (r.acts_as as string | null) ?? null,
      }));
    },
    FALLBACK_ROLES,
  );
}

/** True when `office` is a launcher office (or one of the fallbacks pre-deploy). */
export async function isLauncherOffice(office: string): Promise<boolean> {
  return (await getLauncherOffices()).some((o) => o.name === office);
}

/** name -> display_name map for labels; unknown names render as-is. */
export function labelMap(options: OrgOption[]): Record<string, string> {
  return Object.fromEntries(options.map((o) => [o.name, o.displayName]));
}

/**
 * The role whose in-app permissions this launcher role gets:
 * launcher_permission_role(role) = coalesce(acts_as, name). Falls back to the
 * role itself if the function isn't deployed yet. A role the launcher doesn't
 * know resolves to "" (least privilege). acts_as can never yield admin.
 */
export async function resolvePermRole(
  role: string | null | undefined,
  supabase: SupabaseClient = createAdminClient(),
): Promise<string> {
  if (!role) return "";
  let perm: string | null = role;
  try {
    const { data, error } = await supabase.rpc("launcher_permission_role", { p_role: role });
    if (error) {
      if (!["PGRST202", "42883", "PGRST204"].includes(error.code ?? "")) {
        console.warn(`[launcher-org] launcher_permission_role: ${error.message}`);
      }
      perm = role;
    } else {
      perm = typeof data === "string" ? data : null;
    }
  } catch {
    perm = role;
  }
  if (!perm) return "";
  if (perm === "admin" && role !== "admin") return "";
  return perm;
}
