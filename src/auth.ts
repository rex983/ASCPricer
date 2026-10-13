import NextAuth from "next-auth";
import Google from "next-auth/providers/google";
import Credentials from "next-auth/providers/credentials";
import type { JWT } from "next-auth/jwt";
import { createAdminClient } from "@/lib/supabase/admin";
import { verifyLauncherSsoToken } from "@/lib/sso/verifier";
import type { UserRole, Office } from "@/types/auth";
import { timingSafeEqual } from "crypto";

// Dev bypass ONLY when NODE_ENV is literally "development" — never via env flag
const isDev = process.env.NODE_ENV === "development";

const ADMIN_EMAIL = "rex@bigbuildingsdirect.com";

// Break-glass password login for the bootstrap admin. OFF unless
// ENABLE_ADMIN_PASSWORD_LOGIN=true is set explicitly. Even when on, the
// account must have an active row in profiles — there is no synthetic
// fallback identity.
const adminPasswordLoginEnabled =
  process.env.ENABLE_ADMIN_PASSWORD_LOGIN === "true";

// Re-validate the session against profiles at most this often. Also
// refreshes role/office, replacing the old 5-minute role refresh.
const RECHECK_INTERVAL_MS = 60_000;

type ProfileRow = {
  id: string;
  role: string;
  office: string | null;
  is_active: boolean | null;
  signed_out_at: string | null;
};

const PROFILE_COLUMNS = "id, role, office, is_active, signed_out_at";

/** Timing-safe string comparison to prevent timing attacks on password checks. */
function safeCompare(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  return timingSafeEqual(Buffer.from(a), Buffer.from(b));
}

/** Copy role/office from the profile row onto the token. */
function applyProfile(token: JWT, profile: ProfileRow) {
  token.role = profile.role as UserRole;
  token.profileId = profile.id;
  token.office = profile.office ? (profile.office as Office) : undefined;
}

/**
 * Re-read the profile and decide whether the session is still valid.
 * Returns false to revoke, true to keep. Throws on transient DB errors so
 * the caller can fail open without bumping checkedAt.
 */
async function recheckProfile(token: JWT, email: string): Promise<boolean> {
  const supabase = createAdminClient();
  const { data: profile, error } = await supabase
    .from("profiles")
    .select(PROFILE_COLUMNS)
    .eq("email", email)
    .maybeSingle<ProfileRow>();
  if (error) throw error;
  if (!profile) return false;
  if (profile.is_active === false) return false;
  const signedInAt =
    (token.signedInAt as number | undefined) ?? (token.iat as number | undefined) ?? 0;
  if (profile.signed_out_at && signedInAt > 0) {
    const signedOutSec = Math.floor(new Date(profile.signed_out_at).getTime() / 1000);
    if (signedOutSec > signedInAt) return false;
  }
  applyProfile(token, profile);
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

      // Dev bypass — local development only, never in production
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

      // Dev synthetic user — local development only.
      if (isDev && user.id === "dev-user-001") return true;

      if (account?.provider === "google") {
        if (!user.email.endsWith("@bigbuildingsdirect.com")) return false;

        // Auto-provision a profile row for Google Workspace users if one doesn't exist
        try {
          const supabase = createAdminClient();
          const { data: existing } = await supabase
            .from("profiles")
            .select("id, is_active")
            .eq("email", user.email)
            .single();

          if (!existing) {
            await supabase.from("profiles").insert({
              email: user.email,
              full_name: user.name || user.email.split("@")[0],
              role: "sales_rep",
            });
          } else if (existing.is_active === false) {
            return false;
          }
        } catch {
          // Non-blocking — profile will be created on next sign-in if this fails
        }

        return true;
      }

      // Launcher SSO / admin credentials: the identity is already verified,
      // but the shared profile must exist and be active.
      try {
        const supabase = createAdminClient();
        const { data: profile } = await supabase
          .from("profiles")
          .select("id, is_active")
          .eq("email", user.email)
          .single();
        if (!profile) return false;
        if (profile.is_active === false) return false;
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

        // Dev user gets admin only in development
        if (isDev && user.id === "dev-user-001") {
          token.role = "admin" as UserRole;
          token.profileId = "dev-user-001";
          return token;
        }

        // All users (including admin) — look up actual role from DB
        try {
          const supabase = createAdminClient();
          const { data: profile } = await supabase
            .from("profiles")
            .select(PROFILE_COLUMNS)
            .eq("email", user.email)
            .single<ProfileRow>();

          if (profile) {
            applyProfile(token, profile);
          } else {
            token.role = "sales_rep" as UserRole;
          }
        } catch {
          token.role = "sales_rep" as UserRole;
        }
        return token;
      }

      // Subsequent requests: re-validate at most every 60s. Revokes the
      // session on deactivation / launcher sign-out and refreshes role.
      if (isDev && token.profileId === "dev-user-001") return token;
      const email = token.email as string | undefined;
      if (!email) return null;
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
        session.user.profileId = token.profileId as string;
        if (token.office) session.user.office = token.office as Office;
      }
      return session;
    },
  },
  session: {
    strategy: "jwt",
    maxAge: 8 * 60 * 60, // 8 hours (must stay <= 12h)
  },
});
