import type { DefaultSession } from "next-auth";

export type UserRole = "admin" | "senior_manager" | "sales_rep" | "bst";
export type Office = "Harbor";

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
