import NextAuth from "next-auth";
import Google from "next-auth/providers/google";
import Credentials from "next-auth/providers/credentials";
import { createAdminClient } from "@/lib/supabase/admin";
import type { UserRole, Office } from "@/types/auth";
import { timingSafeEqual } from "crypto";

// Dev bypass ONLY when NODE_ENV is literally "development" — never via env flag
const isDev = process.env.NODE_ENV === "development";

const ADMIN_EMAIL = "rex@bigbuildingsdirect.com";

/** Timing-safe string comparison to prevent timing attacks on password checks. */
function safeCompare(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  return timingSafeEqual(Buffer.from(a), Buffer.from(b));
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

      // Rate limit by email
      if (!checkRateLimit(email.toLowerCase())) {
        console.warn(`Rate limited login attempt for ${email}`);
        return null;
      }

      // Admin login — requires ADMIN_PASSWORD env var (min 8 chars)
      const adminPw = (process.env.ADMIN_PASSWORD || "").trim();
      if (
        email === ADMIN_EMAIL &&
        adminPw.length >= 8 &&
        safeCompare(password, adminPw)
      ) {
        // Resolve the real profile UUID from the database
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
        // Fallback if profile not found (shouldn't happen, but don't lock out admin)
        return { id: "admin-001", email, name: "Rex", image: null };
      }

      // Dev bypass — local development only, never in production
      if (isDev) {
        return {
          id: "dev-user-001",
          email,
          name: email.split("@")[0],
          image: null,
        };
      }

      // No other credential logins allowed — all employees must use Google OAuth
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

      if (account?.provider === "google") {
        if (!user.email.endsWith("@bigbuildingsdirect.com")) return false;

        // Auto-provision a profile row for Google Workspace users if one doesn't exist
        try {
          const supabase = createAdminClient();
          const { data: existing } = await supabase
            .from("profiles")
            .select("id")
            .eq("email", user.email)
            .single();

          if (!existing) {
            await supabase.from("profiles").insert({
              email: user.email,
              full_name: user.name || user.email.split("@")[0],
              role: "sales_rep",
            });
          }
        } catch {
          // Non-blocking — profile will be created on next sign-in if this fails
        }

        return true;
      }

      // Credentials — already validated in authorize()
      return true;
    },
    async jwt({ token, user }) {
      const email = user?.email || (token.email as string);

      // On initial sign-in, populate token from DB
      if (user?.email) {
        // Dev user gets admin only in development
        if (isDev && user.id === "dev-user-001") {
          token.role = "admin" as UserRole;
          token.profileId = "dev-user-001";
          token.roleRefreshedAt = Date.now();
          return token;
        }

        // All users (including admin) — look up actual role from DB
        try {
          const supabase = createAdminClient();
          const { data: profile } = await supabase
            .from("profiles")
            .select("id, role, office")
            .eq("email", user.email)
            .single();

          if (profile) {
            token.role = profile.role as UserRole;
            token.profileId = profile.id;
            if (profile.office) token.office = profile.office as Office;
          } else {
            token.role = "sales_rep" as UserRole;
          }
        } catch {
          token.role = "sales_rep" as UserRole;
        }
        token.roleRefreshedAt = Date.now();
        return token;
      }

      // On subsequent requests, refresh role from DB every 5 minutes
      const lastRefresh = (token.roleRefreshedAt as number) || 0;
      const REFRESH_INTERVAL = 5 * 60 * 1000; // 5 minutes
      if (email && Date.now() - lastRefresh > REFRESH_INTERVAL) {
        try {
          const supabase = createAdminClient();
          const { data: profile } = await supabase
            .from("profiles")
            .select("id, role, office")
            .eq("email", email)
            .single();

          if (profile) {
            token.role = profile.role as UserRole;
            token.profileId = profile.id;
            token.office = profile.office ? (profile.office as Office) : undefined;
          }
        } catch {
          // Keep existing token values on refresh failure
        }
        token.roleRefreshedAt = Date.now();
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
    maxAge: 8 * 60 * 60, // 8 hours
  },
});
