import NextAuth from "next-auth";
import Google from "next-auth/providers/google";
import Credentials from "next-auth/providers/credentials";
import { createAdminClient } from "@/lib/supabase/admin";
import type { UserRole } from "@/types/auth";

const isDev =
  process.env.NODE_ENV === "development" ||
  process.env.AUTH_DEV_BYPASS === "true";

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

      // Admin hardcoded login
      if (
        email === "rex@bigbuildingsdirect.com" &&
        password === process.env.ADMIN_PASSWORD
      ) {
        return { id: "admin-001", email, name: "Rex", image: null };
      }

      // Dev bypass
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
          .select("id, email, name, role")
          .eq("email", email)
          .single();

        if (!profile) return null;

        return {
          id: profile.id,
          email: profile.email,
          name: profile.name || null,
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
        return true;
      }

      // Credentials — already validated in authorize()
      return true;
    },
    async jwt({ token, user }) {
      if (user?.email) {
        // Admin or dev user
        if (user.email === "rex@bigbuildingsdirect.com" || isDev) {
          token.role = "admin" as UserRole;
          token.profileId = user.id || "admin-001";
          return token;
        }

        // DB user
        try {
          const supabase = createAdminClient();
          const { data: profile } = await supabase
            .from("profiles")
            .select("id, role")
            .eq("email", user.email)
            .single();

          if (profile) {
            token.role = profile.role as UserRole;
            token.profileId = profile.id;
          }
        } catch {
          // fallback
          token.role = "sales_rep" as UserRole;
        }
      }
      return token;
    },
    async session({ session, token }) {
      if (token) {
        session.user.role = token.role as UserRole;
        session.user.profileId = token.profileId as string;
      }
      return session;
    },
  },
  session: {
    strategy: "jwt",
  },
});
