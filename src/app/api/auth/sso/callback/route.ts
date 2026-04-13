import { NextRequest, NextResponse } from "next/server";
import { createRemoteJWKSet, jwtVerify } from "jose";
import { createAdminClient } from "@/lib/supabase/admin";

let jwks: ReturnType<typeof createRemoteJWKSet> | null = null;

function getJwks() {
  if (jwks) return jwks;
  const launcherUrl = process.env.BBD_LAUNCHER_URL;
  if (!launcherUrl) throw new Error("BBD_LAUNCHER_URL is not configured");
  jwks = createRemoteJWKSet(new URL(`${launcherUrl}/api/sso/jwks`));
  return jwks;
}

/**
 * GET /api/auth/sso/callback?sso_token=<jwt>
 *
 * BBDLauncher redirects here with a signed JWT. We verify it,
 * upsert the profile, then redirect to the NextAuth credentials
 * sign-in which picks up the profile from the DB.
 */
export async function GET(req: NextRequest) {
  const token = req.nextUrl.searchParams.get("sso_token");
  if (!token) {
    return NextResponse.redirect(new URL("/login?error=missing_token", req.url));
  }

  try {
    const { payload } = await jwtVerify(token, getJwks(), {
      issuer: "bbd-launcher",
      audience: "asc-pricing",
      maxTokenAge: 120, // allow a little clock skew beyond the 60s expiry
    });

    const email = payload.email as string;
    const name = (payload.name as string) || email.split("@")[0];
    const role = payload.role as string | undefined;
    const profileId = payload.profile_id as string | undefined;

    if (!email) {
      return NextResponse.redirect(new URL("/login?error=invalid_token", req.url));
    }

    // Upsert profile — create if not exists, update name if it does
    const supabase = createAdminClient();
    const { data: existing } = await supabase
      .from("profiles")
      .select("id, role")
      .eq("email", email)
      .maybeSingle();

    if (!existing) {
      await supabase.from("profiles").insert({
        id: profileId || undefined,
        email,
        name,
        role: role || "viewer",
      });
    }

    // Redirect to NextAuth sign-in with credentials provider
    // The credentials provider will look up the profile by email
    const signInUrl = new URL("/api/auth/callback/bbd-sso", req.url);
    signInUrl.searchParams.set("email", email);
    signInUrl.searchParams.set("callbackUrl", "/calculator");

    // Use a POST form auto-submit to trigger the credentials callback
    const html = `<!DOCTYPE html>
<html><body>
<form id="f" method="POST" action="${signInUrl.pathname}">
  <input type="hidden" name="email" value="${escapeHtml(email)}" />
  <input type="hidden" name="callbackUrl" value="/calculator" />
</form>
<script>document.getElementById('f').submit();</script>
</body></html>`;

    return new NextResponse(html, {
      headers: { "Content-Type": "text/html" },
    });
  } catch (err) {
    console.error("SSO token verification failed:", err);
    return NextResponse.redirect(new URL("/login?error=sso_failed", req.url));
  }
}

function escapeHtml(str: string): string {
  return str.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
