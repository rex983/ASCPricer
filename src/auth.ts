import NextAuth from "next-auth";
import Google from "next-auth/providers/google";
import Credentials from "next-auth/providers/credentials";
import type { JWT } from "next-auth/jwt";
import type { SupabaseClient } from "@supabase/supabase-js";
import { createAdminClient } from "@/lib/supabase/admin";
import { verifyLauncherSsoToken } from "@/lib/sso/verifier";
import {
  ACCESS_GRACE_MS,
  LAUNCHER_CHECK_FAILED,
  NO_LAUNCHER_ACCESS,
  launcherAccess,
} from "@/lib/launcher-access";
import { resolvePermRole } from "@/lib/launcher-org";
import type { UserRole, Office, ViewAs } from "@/types/auth";
import { timingSafeEqual } from "crypto";

// Dev bypass ONLY when NODE_ENV is literally "development" — never via env flag
const isDev = process.env.NODE_ENV === "development";

const ADMIN_EMAIL = "rex@bigbuildingsdirect.com";

// Break-glass password login for the bootstrap admin. OFF unless
// ENABLE_ADMIN_PASSWORD_LOGIN=true is set explicitly. It only authenticates:
// the account still needs an active profiles row AND launcher access to this
// app (launcher_can_open_app === true), like every other sign-in.
const adminPasswordLoginEnabled =
  process.env.ENABLE_ADMIN_PASSWORD_LOGIN === "true";

// Re-validate the session against profiles (and launcher app access) at
// most this often. Also refreshes role/office/department/is_it/name, so a
// role change in the launcher lands within a minute.
const RECHECK_INTERVAL_MS = 60_000;

// Hard cap on a launcher admin "View as" session.
const VIEW_AS_MAX_MS = 60 * 60 * 1000;

type ProfileRow = {
  id: string;
  email: string;
  full_name: string | null;
  role: string;
  office: string | null;
  department: string | null;
  is_it: boolean | null;
  is_active: boolean | null;
  signed_out_at: string | null;
};

const PROFILE_COLUMNS =
  "id, email, full_name, role, office, department, is_it, is_active, signed_out_at";

/** Timing-safe string comparison to prevent timing attacks on password checks. */
function safeCompare(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  return timingSafeEqual(Buffer.from(a), Buffer.from(b));
}

async function loadProfile(
  supabase: SupabaseClient,
  column: "email" | "id",
  value: string,
): Promise<ProfileRow | null> {
  const { data, error } = await supabase
    .from("profiles")
    .select(PROFILE_COLUMNS)
    .eq(column, value)
    .maybeSingle<ProfileRow>();
  if (error) throw error;
  return data;
}

/** The launcher admin behind a "View as" session must still be an active admin. */
function isActiveAdmin(profile: ProfileRow | null): boolean {
  return !!profile && profile.is_active !== false && profile.role === "admin";
}

/** Copy the launcher-owned identity fields from the profile row onto the token. */
function applyProfile(token: JWT, profile: ProfileRow, permRole: string) {
  token.role = profile.role as UserRole;
  // In-app permissions use permRole = launcher_permission_role(role)
  // (custom roles inherit their acts_as built-in role); display uses role.
  token.permRole = permRole as UserRole;
  token.profileId = profile.id;
  token.office = profile.office ? (profile.office as Office) : undefined;
  token.department = profile.department ?? null;
  token.isIt = profile.is_it === true;
  if (profile.full_name) token.name = profile.full_name;
}

/**
 * Re-read the profile (+ launcher access, + permission role, + the view-as
 * actor) and decide whether the session is still valid. Returns false to
 * revoke, true to keep. Throws on transient DB errors so the caller keeps the
 * session without bumping checkedAt.
 *
 * Launcher access is fail-closed: false/NULL revokes. If the RPC is missing or
 * erroring, the session survives only ACCESS_GRACE_MS past its last
 * successful access check.
 */
async function recheckProfile(token: JWT, email: string): Promise<boolean> {
  const supabase = createAdminClient();
  const viewAs = (token.viewAs as ViewAs | null | undefined) ?? null;
  const profileId = token.profileId as string | undefined;
  const tokenRole = token.role as string | undefined;

  const [profile, knownAccess, actor, knownPerm] = await Promise.all([
    loadProfile(supabase, "email", email),
    profileId ? launcherAccess(profileId, supabase) : Promise.resolve(null),
    viewAs
      ? loadProfile(supabase, "id", viewAs.actorProfileId)
      : Promise.resolve(null),
    resolvePermRole(tokenRole, supabase),
  ]);

  if (!profile) return false;
  if (profile.is_active === false) return false;
  const signedInAt =
    (token.signedInAt as number | undefined) ?? (token.iat as number | undefined) ?? 0;
  if (profile.signed_out_at && signedInAt > 0) {
    const signedOutSec = Math.floor(new Date(profile.signed_out_at).getTime() / 1000);
    if (signedOutSec > signedInAt) return false;
  }

  const access =
    knownAccess !== null && profileId === profile.id
      ? knownAccess
      : await launcherAccess(profile.id, supabase);
  if (access === false) return false;
  if (access === true) {
    token.accessOkAt = Date.now();
  } else {
    const okAt = (token.accessOkAt as number | undefined) ?? 0;
    if (Date.now() - okAt > ACCESS_GRACE_MS) return false;
  }

  if (viewAs && !isActiveAdmin(actor)) return false;
  const permRole =
    profile.role === tokenRole ? knownPerm : await resolvePermRole(profile.role, supabase);
  applyProfile(token, profile, permRole);
  return true;
}

// Only include Google provider if credentials are configured
const providers = [];

if (process.env.AUTH_GOOGLE_ID && process.env.AUTH_GOOGLE_SECRET) {
  providers.push(
    Google({
      clientId: process.env.AUTH_GOOGLE_ID,
      clientSecret: process.env.AUTH_GOOGLE_SECRET,
      authorization: {
        params: {
          hd: "bigbuildingsdirect.com",
          prompt: "select_account",
        },
      },
    })
  );
}

// Launcher JWT-SSO: the launcher redirects to /sso/callback?sso_token=...,
// which posts the token here. The token is verified (ES256, issuer
// bbd-launcher, audience asc-pricer, single-use jti) before we trust it.
providers.push(
  Credentials({
    id: "launcher-sso",
    name: "Launcher SSO",
    credentials: {
      token: { label: "SSO Token", type: "text" },
    },
    async authorize(credentials) {
      const token = credentials?.token as string;
      if (!token) return null;
      try {
        const claims = await verifyLauncherSsoToken(token);
        return {
          id: claims.profile_id || claims.sub,
          email: claims.email,
          name: claims.name || null,
          image: null,
          act: claims.act,
        };
      } catch (err) {
        console.error("Launcher SSO token verification failed:", err);
        return null;
      }
    },
  })
);

// In-memory rate limiter for credential login attempts
const loginAttempts = new Map<string, { count: number; resetAt: number }>();
const MAX_ATTEMPTS = 5;
const WINDOW_MS = 60_000; // 1 minute

function checkRateLimit(key: string): boolean {
  const now = Date.now();
  const entry = loginAttempts.get(key);
  if (!entry || now > entry.resetAt) {
    loginAttempts.set(key, { count: 1, resetAt: now + WINDOW_MS });
    return true;
  }
  entry.count++;
  return entry.count <= MAX_ATTEMPTS;
}

providers.push(
  Credentials({
    id: "credentials",
    name: "Email & Password",
    credentials: {
      email: { label: "Email", type: "email" },
      password: { label: "Password", type: "password" },
    },
    async authorize(credentials) {
      const email = credentials?.email as string;
      const password = credentials?.password as string;
      if (!email || !password) return null;

      // Dev password bypass — local development only, never in production.
      // It only authenticates: signIn still requires an active profile for
      // this email and launcher access (launcher_can_open_app === true).
      if (isDev) {
        return {
          id: "dev-user-001",
          email,
          name: email.split("@")[0],
          image: null,
        };
      }

      if (!adminPasswordLoginEnabled) return null;

      // Rate limit by email
      if (!checkRateLimit(email.toLowerCase())) {
        console.warn(`Rate limited login attempt for ${email}`);
        return null;
      }

      // Admin login — requires ENABLE_ADMIN_PASSWORD_LOGIN=true and
      // ADMIN_PASSWORD (min 8 chars)
      const adminPw = (process.env.ADMIN_PASSWORD || "").trim();
      if (
        email === ADMIN_EMAIL &&
        adminPw.length >= 8 &&
        safeCompare(password, adminPw)
      ) {
        // Resolve the real profile UUID from the database. No profile, no
        // login — the old synthetic "admin-001" fallback is gone.
        try {
          const supabase = createAdminClient();
          const { data: profile } = await supabase
            .from("profiles")
            .select("id, full_name")
            .eq("email", ADMIN_EMAIL)
            .single();

          if (profile) {
            return { id: profile.id, email, name: profile.full_name || "Rex", image: null };
          }
        } catch {
          // Fall through — profile lookup failed
        }
        return null;
      }

      // No other credential logins allowed — all employees must use Google
      // OAuth or the launcher.
      return null;
    },
  })
);

export const { handlers, signIn, signOut, auth } = NextAuth({
  trustHost: true,
  providers,
  pages: {
    signIn: "/login",
    error: "/login",
  },
  callbacks: {
    async signIn({ user, account }) {
      if (!user.email) return false;

      if (
        account?.provider === "google" &&
        !user.email.endsWith("@bigbuildingsdirect.com")
      ) {
        return false;
      }

      // Every provider (Google, launcher SSO, break-glass password, dev
      // password): the launcher's profiles row is the only identity source,
      // and launcher_can_open_app must return exactly true. No bypasses.
      // (Profiles are never created here — the launcher's Admin -> Users is
      // the only place users are added.)
      try {
        const supabase = createAdminClient();
        const profile = await loadProfile(supabase, "email", user.email);
        if (!profile) return false;
        if (profile.is_active === false) return false;

        const access = await launcherAccess(profile.id, supabase);
        if (access === "error") return `/login?error=${LAUNCHER_CHECK_FAILED}`;
        if (access !== true) return `/login?error=${NO_LAUNCHER_ACCESS}`;

        // Launcher admin "View as": the actor must be an active admin.
        if (user.act) {
          if (account?.provider !== "launcher-sso") return false;
          const actor = await loadProfile(supabase, "id", user.act.profile_id);
          if (!actor || !isActiveAdmin(actor)) return false;
          if (actor.email.toLowerCase() !== user.act.sub.toLowerCase()) {
            return false;
          }
        }
        return true;
      } catch {
        return false;
      }
    },
    async jwt({ token, user }) {
      // On initial sign-in, populate token from DB
      if (user?.email) {
        token.signedInAt = Math.floor(Date.now() / 1000);
        token.checkedAt = Date.now();
        // signIn just confirmed launcher access.
        token.accessOkAt = Date.now();
        token.viewAs = user.act
          ? ({
              actorProfileId: user.act.profile_id,
              actorEmail: user.act.sub,
              actorName: user.act.name || user.act.sub,
              startedAt: Date.now(),
            } satisfies ViewAs)
          : null;

        // All users (including admin) — look up actual role from DB. signIn
        // just verified the profile; if it can't be re-read, refuse.
        try {
          const supabase = createAdminClient();
          const profile = await loadProfile(supabase, "email", user.email);
          if (!profile) return null;
          applyProfile(token, profile, await resolvePermRole(profile.role, supabase));
        } catch {
          return null;
        }
        return token;
      }

      // Subsequent requests: re-validate at most every 60s. Revokes the
      // session on deactivation / launcher sign-out / lost launcher access /
      // view-as actor no longer an admin, and refreshes role + org fields.
      const email = token.email as string | undefined;
      if (!email) return null;
      const viewAs = (token.viewAs as ViewAs | null | undefined) ?? null;
      if (viewAs && Date.now() - viewAs.startedAt > VIEW_AS_MAX_MS) return null;
      const checkedAt = (token.checkedAt as number | undefined) ?? 0;
      if (Date.now() - checkedAt < RECHECK_INTERVAL_MS) return token;
      try {
        if (!(await recheckProfile(token, email))) return null;
        token.checkedAt = Date.now();
      } catch (err) {
        // Transient DB error — keep the session, retry on the next request.
        console.error("Session re-validation failed:", err);
      }
      return token;
    },
    async session({ session, token }) {
      if (token) {
        session.user.role = token.role as UserRole;
        session.user.permRole = (token.permRole as UserRole | undefined) ?? "";
        session.user.profileId = token.profileId as string;
        if (token.office) session.user.office = token.office as Office;
        session.user.department = (token.department as string | null | undefined) ?? null;
        session.user.isIt = token.isIt === true;
        if (typeof token.name === "string") session.user.name = token.name;
      }
      session.viewAs = (token?.viewAs as ViewAs | null | undefined) ?? null;
      return session;
    },
  },
  session: {
    strategy: "jwt",
    maxAge: 8 * 60 * 60, // 8 hours (must stay <= 12h)
  },
});
