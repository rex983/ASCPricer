"use client";

import { Suspense, useEffect, useState } from "react";
import { useRouter, useSearchParams } from "next/navigation";
import { signIn } from "next-auth/react";

function Callback() {
  const router = useRouter();
  const searchParams = useSearchParams();
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const token = searchParams.get("sso_token");
    if (!token) {
      setError("Missing SSO token");
      return;
    }
    (async () => {
      const res = await signIn("launcher-sso", {
        token,
        callbackUrl: "/calculator",
        redirect: false,
      });
      if (res?.ok && res.url) {
        router.replace(res.url);
      } else {
        setError("Sign-in failed. Token may be expired or invalid.");
      }
    })();
  }, [router, searchParams]);

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
