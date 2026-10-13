"use client";

import { Suspense, useEffect, useRef, useState } from "react";
import { useRouter, useSearchParams } from "next/navigation";
import { signIn } from "next-auth/react";

function Callback() {
  const router = useRouter();
  const searchParams = useSearchParams();
  const token = searchParams.get("sso_token");
  const [failed, setFailed] = useState(false);
  // Launcher tokens are single-use (jti replay protection), so make sure we
  // only redeem once even if the effect runs twice (React StrictMode).
  const redeemed = useRef(false);

  useEffect(() => {
    if (!token || redeemed.current) return;
    redeemed.current = true;
    (async () => {
      const res = await signIn("launcher-sso", {
        token,
        callbackUrl: "/calculator",
        redirect: false,
      });
      if (res?.ok && res.url) {
        router.replace(res.url);
      } else {
        setFailed(true);
      }
    })();
  }, [router, token]);

  const error = !token
    ? "Missing SSO token"
    : failed
      ? "Sign-in failed. Token may be expired or invalid."
      : null;

  return (
    <div className="flex min-h-screen items-center justify-center">
      <div className="text-center space-y-2">
        {error ? (
          <>
            <p className="text-destructive font-medium">{error}</p>
            <a href="/login" className="text-sm text-muted-foreground underline">
              Go to login
            </a>
          </>
        ) : (
          <p className="text-muted-foreground">Signing you in…</p>
        )}
      </div>
    </div>
  );
}

export default function SsoCallbackPage() {
  return (
    <Suspense>
      <Callback />
    </Suspense>
  );
}
