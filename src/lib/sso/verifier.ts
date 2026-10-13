import { createRemoteJWKSet, jwtVerify } from "jose";
import { createAdminClient } from "@/lib/supabase/admin";

// Public JWKS endpoint published by the BBD Launcher (its IdP).
// Set LAUNCHER_JWKS_URL on Vercel; defaults to the production deployment.
const JWKS_URL =
  process.env.LAUNCHER_JWKS_URL ||
  "https://bbd-launcher.vercel.app/api/sso/jwks";

const JWKS = createRemoteJWKSet(new URL(JWKS_URL));

// Must match launcher_sso_configs.jwt_audience for this app.
const AUDIENCE = "asc-pricer";

export interface LauncherSsoClaims {
  sub: string;
  email: string;
  name: string;
  role: string;
  profile_id: string;
}

/**
 * Record the token's jti so each launcher token can be redeemed only once.
 * Unique violation => replay, reject. Missing table (migration not applied
 * yet) => warn and allow so a deploy ahead of the launcher migration doesn't
 * break sign-in. Any other DB error fails closed.
 */
async function consumeJti(jti: string, expSec: number): Promise<void> {
  const supabase = createAdminClient();
  const { error } = await supabase.from("sso_consumed_jti").insert({
    jti,
    audience: AUDIENCE,
    expires_at: new Date(expSec * 1000).toISOString(),
  });
  if (!error) return;
  if (error.code === "23505") {
    throw new Error("SSO token has already been used");
  }
  if (error.code === "42P01" || error.code === "PGRST205") {
    console.warn(
      "sso_consumed_jti table missing — SSO replay protection disabled until the launcher migration is applied",
    );
    return;
  }
  throw new Error(`Failed to record SSO token: ${error.message}`);
}

export async function verifyLauncherSsoToken(
  token: string,
): Promise<LauncherSsoClaims> {
  const { payload } = await jwtVerify(token, JWKS, {
    algorithms: ["ES256"],
    issuer: "bbd-launcher",
    audience: AUDIENCE,
    maxTokenAge: "120s",
    clockTolerance: "10s",
  });

  if (typeof payload.jti !== "string" || payload.jti.length === 0) {
    throw new Error("SSO token missing jti claim");
  }
  if (typeof payload.email !== "string") {
    throw new Error("SSO token missing email claim");
  }

  // maxTokenAge guarantees iat is present; exp is set by the launcher (60s).
  const expSec =
    typeof payload.exp === "number" ? payload.exp : (payload.iat as number) + 120;
  await consumeJti(payload.jti, expSec);

  return {
    sub: String(payload.sub ?? payload.email),
    email: payload.email,
    name: typeof payload.name === "string" ? payload.name : "",
    role: typeof payload.role === "string" ? payload.role : "",
    profile_id:
      typeof payload.profile_id === "string"
        ? payload.profile_id
        : String(payload.sub ?? ""),
  };
}
