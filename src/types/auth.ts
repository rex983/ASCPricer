import type { DefaultSession } from "next-auth";

export type UserRole = "admin" | "sales_rep" | "viewer";

declare module "next-auth" {
  interface Session {
    user: {
      role: UserRole;
      profileId: string;
    } & DefaultSession["user"];
  }

  interface JWT {
    role?: UserRole;
    profileId?: string;
  }
}
