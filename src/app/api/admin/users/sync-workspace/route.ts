import { NextResponse } from "next/server";
import { auth } from "@/auth";
import { createAdminClient } from "@/lib/supabase/admin";
import { logAudit } from "@/lib/audit";
import { SignJWT, importPKCS8 } from "jose";

const SCOPES = "https://www.googleapis.com/auth/admin.directory.user.readonly";
const TOKEN_URL = "https://oauth2.googleapis.com/token";
const DIRECTORY_URL = "https://admin.googleapis.com/admin/directory/v1/users";

/** Exchange a service-account JWT for a Google access token */
async function getAccessToken(): Promise<string> {
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const rawKey = process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY;
  const adminEmail = process.env.GOOGLE_ADMIN_EMAIL;

  if (!email || !rawKey || !adminEmail) {
    throw new Error(
      "Missing env vars: GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY, GOOGLE_ADMIN_EMAIL"
    );
  }

  // The key comes from Vercel env with literal \n — convert to real newlines
  const privateKey = rawKey.replace(/\\n/g, "\n");
  const key = await importPKCS8(privateKey, "RS256");

  const now = Math.floor(Date.now() / 1000);
  const jwt = await new SignJWT({
    iss: email,
    sub: adminEmail, // impersonate a workspace admin
    scope: SCOPES,
    aud: TOKEN_URL,
  })
    .setProtectedHeader({ alg: "RS256", typ: "JWT" })
    .setIssuedAt(now)
    .setExpirationTime(now + 3600)
    .sign(key);

  const res = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt,
    }),
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Google token exchange failed: ${err}`);
  }

  const data = await res.json();
  return data.access_token;
}

interface GoogleUser {
  primaryEmail: string;
  name: { fullName: string };
  suspended: boolean;
  orgUnitPath: string;
}

/** Fetch all active users from Google Workspace */
async function listWorkspaceUsers(accessToken: string): Promise<GoogleUser[]> {
  const domain = "bigbuildingsdirect.com";
  const allUsers: GoogleUser[] = [];
  let pageToken: string | undefined;

  do {
    const params = new URLSearchParams({
      domain,
      maxResults: "500",
      orderBy: "email",
      projection: "basic",
    });
    if (pageToken) params.set("pageToken", pageToken);

    const res = await fetch(`${DIRECTORY_URL}?${params}`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });

    if (!res.ok) {
      const err = await res.text();
      throw new Error(`Directory API failed: ${err}`);
    }

    const data = await res.json();
    if (data.users) {
      allUsers.push(...data.users);
    }
    pageToken = data.nextPageToken;
  } while (pageToken);

  // Only return non-suspended users
  return allUsers.filter((u) => !u.suspended);
}

/** POST /api/admin/users/sync-workspace — pull Google Workspace users into profiles */
export async function POST() {
  const session = await auth();
  if (!session?.user || session.user.role !== "admin") {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  try {
    const accessToken = await getAccessToken();
    const workspaceUsers = await listWorkspaceUsers(accessToken);

    const supabase = createAdminClient();

    // Fetch existing profiles by email for fast lookup
    const { data: existing } = await supabase.from("profiles").select("id, email");
    const existingEmails = new Set((existing ?? []).map((p) => p.email.toLowerCase()));

    let created = 0;
    let skipped = 0;
    const newUsers: { email: string; name: string }[] = [];

    for (const gu of workspaceUsers) {
      const email = gu.primaryEmail.toLowerCase();

      if (existingEmails.has(email)) {
        skipped++;
        continue;
      }

      const { error } = await supabase.from("profiles").insert({
        email,
        name: gu.name.fullName,
        role: "viewer", // default — admin can promote later
      });

      if (!error) {
        created++;
        newUsers.push({ email, name: gu.name.fullName });
        existingEmails.add(email); // prevent duplicates in same batch
      } else {
        // Likely a race condition duplicate — skip
        skipped++;
      }
    }

    await logAudit({
      userId: session.user.profileId,
      userEmail: session.user.email,
      action: "sync_workspace",
      resourceType: "profile",
      details: {
        workspace_total: workspaceUsers.length,
        created,
        skipped,
        new_users: newUsers.slice(0, 20), // cap detail size
      },
    });

    return NextResponse.json({
      total: workspaceUsers.length,
      created,
      skipped,
      new_users: newUsers,
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unknown error";
    console.error("Workspace sync failed:", message);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
