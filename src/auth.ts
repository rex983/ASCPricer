import NextAuth from "next-auth";
import Google from "next-auth/providers/google";
import Credentials from "next-auth/providers/credentials";
import { createAdminClient } from "@/lib/supabase/admin";
import type { UserRole, Office } from "@/types/auth";

// Dev bypass ONLY when NODE_ENV is literally "development" — never via env flag
const isDev = process.env.NODE_ENV === "development";

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

      // Admin login — requires ADMIN_PASSWORD env var (min 8 chars)
      const adminPw = (process.env.ADMIN_PASSWORD || "").trim();
      if (
        email === "rex@bigbuildingsdirect.com" &&
        adminPw.length >= 8 &&
        password === adminPw
      ) {
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

      // DB lookup
      try {
        const supabase = createAdminClient();
        const { data: profile } = await supabase
          .from("profiles")
          .select("id, email, full_name, role")
          .eq("email", email)
          .single();

        if (!profile) return null;

        return {
          id: profile.id,
          email: profile.email,
          name: profile.full_name || null,
          image: null,
        };
      } catch {
        return null;
      }
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
      if (user?.email) {
        // Hardcoded admin account
        if (user.email === "rex@bigbuildingsdirect.com" && user.id === "admin-001") {
          token.role = "admin" as UserRole;
          token.profileId = "admin-001";
          return token;
        }

        // Dev user gets admin only in development
        if (isDev && user.id === "dev-user-001") {
          token.role = "admin" as UserRole;
          token.profileId = "dev-user-001";
          return token;
        }

        // DB user — look up actual role
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
            // No profile found — default to most restrictive role
            token.role = "sales_rep" as UserRole;
          }
        } catch {
          token.role = "sales_rep" as UserRole;
        }
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
