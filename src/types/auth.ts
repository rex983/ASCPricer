import type { DefaultSession } from "next-auth";

export type UserRole = "admin" | "manager" | "sales_rep" | "viewer";
export type Office = "Harbor" | "Marion";

declare module "next-auth" {
  interface Session {
    user: {
      role: UserRole;
      profileId: string;
      office?: Office;
    } & DefaultSession["user"];
  }

  interface JWT {
    role?: UserRole;
    profileId?: string;
    office?: Office;
  }
}
