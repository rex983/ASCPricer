import { NextRequest, NextResponse } from "next/server";
import { getToken } from "next-auth/jwt";

const publicPaths = ["/login", "/api/auth", "/sso/callback"];

const SAFE_METHODS = new Set(["GET", "HEAD", "OPTIONS"]);

// Non-GET requests still allowed during a launcher admin "View as" session.
// Only sign-in / sign-out — every other POST/PUT/PATCH/DELETE in this app
// (API routes and any server action, which are POSTs with a Next-Action
// header) writes data.
const VIEW_AS_WRITE_ALLOWLIST = ["/api/auth/"];

const SECURE_COOKIE = "__Secure-authjs.session-token";
const PLAIN_COOKIE = "authjs.session-token";

/** Decode the session JWT (edge-safe) to see whether this is a View-as session. */
async function readViewAs(
  req: NextRequest,
): Promise<{ name: string | null } | null> {
  const secret = process.env.AUTH_SECRET || process.env.NEXTAUTH_SECRET;
  if (!secret) return null;
  const secureCookie = req.cookies.has(SECURE_COOKIE);
  try {
    const token = await getToken({
      req,
      secret,
      secureCookie,
      cookieName: secureCookie ? SECURE_COOKIE : PLAIN_COOKIE,
    });
    if (!token?.viewAs) return null;
    return { name: typeof token.name === "string" ? token.name : null };
  } catch {
    // Undecodable cookie: the route handler's auth() rejects it anyway, and
    // handlers carry their own rejectIfViewingAs() check.
    return null;
  }
}

export async function middleware(req: NextRequest) {
  const { pathname } = req.nextUrl;

  // Allow public paths
  if (publicPaths.some((p) => pathname.startsWith(p))) {
    return NextResponse.next();
  }

  // Check for NextAuth session token cookie
  const token =
    req.cookies.get(SECURE_COOKIE)?.value || req.cookies.get(PLAIN_COOKIE)?.value;

  if (!token) {
    const callback = pathname === "/" ? "/calculator" : pathname;
    // Auto-redirect to Google OAuth instead of showing login page
    const signInUrl = new URL("/api/auth/signin/google", req.url);
    signInUrl.searchParams.set("callbackUrl", callback);
    return NextResponse.redirect(signInUrl);
  }

  // Launcher admin "View as" is read-only: refuse every mutation centrally.
  if (
    !SAFE_METHODS.has(req.method) &&
    !VIEW_AS_WRITE_ALLOWLIST.some((p) => pathname.startsWith(p))
  ) {
    const viewAs = await readViewAs(req);
    if (viewAs) {
      return NextResponse.json(
        { error: `Read-only: viewing as ${viewAs.name || "another user"}` },
        { status: 403 },
      );
    }
  }

  return NextResponse.next();
}

export const config = {
  matcher: [
    "/((?!_next/static|_next/image|favicon.ico|.*\\.(?:svg|png|jpg|jpeg|gif|webp)$).*)",
  ],
};
