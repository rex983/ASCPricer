import NextAuth from "next-auth";
import Google from "next-auth/providers/google";
import Credentials from "next-auth/providers/credentials";
import { createAdminClient } from "@/lib/supabase/admin";
import type { UserRole } from "@/types/auth";

const isDev =
  process.env.NODE_ENV === "development" ||
  process.env.AUTH_DEV_BYPASS === "true";

export const { handlers, signIn, signOut, auth } = NextAuth({
  providers: [
    Google({
      clientId: process.env.AUTH_GOOGLE_ID,
      clientSecret: process.env.AUTH_GOOGLE_SECRET,
      authorization: {
        params: {
          hd: "bigbuildingsdirect.com",
          prompt: "select_account",
        },
      },
    }),
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

        if (
          email === "rex@bigbuildingsdirect.com" &&
          password === process.env.ADMIN_PASSWORD
        ) {
          return { id: "admin-001", email, name: "Rex", image: null };
        }

        if (isDev) {
          return {
            id: "dev-user-001",
            email,
            name: email.split("@")[0],
            image: null,
          };
        }

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
      },
    }),
  ],
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

      if (user.email === "rex@bigbuildingsdirect.com") return true;
      if (isDev) return true;

      try {
        const supabase = createAdminClient();
        const { data: profile } = await supabase
          .from("profiles")
          .select("id")
          .eq("email", user.email)
          .single();
        if (!profile) return false;
      } catch {
        return false;
      }

      return true;
    },
    async jwt({ token, user }) {
      if (user?.email) {
        if (user.email === "rex@bigbuildingsdirect.com" || isDev) {
          token.role = (token.role as UserRole) || "admin";
          token.profileId =
            (token.profileId as string) || user.id || "admin-001";
          return token;
        }

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
