// office-login: checks an Office OS name + PIN and hands the browser a Supabase Auth session.
//
// The browser only holds the public (publishable) key, which RLS blocks from every table.
// On a correct PIN this function makes sure a matching Supabase Auth user exists (tagged with
// app_metadata.app_user_id, which the RLS policies require), then returns a one-time magic-link
// token hash. The browser exchanges it with supabase.auth.verifyOtp() for a normal session.
// No email is ever sent — generateLink only creates the token.
//
// Deploy with verify_jwt = false: callers are not logged in yet, and the publishable key
// is not a JWT. This function does its own check (the PIN).

import { createClient } from "npm:@supabase/supabase-js@2";

const MAX_FAILURES = 10;          // wrong PINs allowed per name...
const LOCKOUT_MINUTES = 15;       // ...within this window

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function secretKey(): string {
  // Prefer the new secret key; fall back to the legacy service_role key until it is disabled.
  // Use the key named "default" if there is one, otherwise any secret key the project has.
  try {
    const keys = JSON.parse(Deno.env.get("SUPABASE_SECRET_KEYS") ?? "{}");
    const key = keys.default ?? Object.values(keys)[0];
    if (key) return key as string;
  } catch (_) { /* fall through */ }
  return Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
}

const admin = createClient(Deno.env.get("SUPABASE_URL")!, secretKey(), {
  auth: { persistSession: false, autoRefreshToken: false },
});

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  if (req.method !== "POST") return json({ error: "Method not allowed" }, 405);

  let name = "", pin = "";
  try {
    const body = await req.json();
    name = String(body?.name ?? "").trim();
    pin = String(body?.pin ?? "");
  } catch (_) {
    return json({ error: "Invalid request" }, 400);
  }
  if (!name || !pin || name.length > 100 || pin.length > 32) {
    return json({ error: "Invalid Name or PIN." }, 401);
  }

  const nameKey = name.toLowerCase();
  const since = new Date(Date.now() - LOCKOUT_MINUTES * 60_000).toISOString();

  const { count: failures, error: countErr } = await admin.from("login_attempts")
    .select("id", { count: "exact", head: true })
    .eq("name_key", nameKey)
    .gte("attempted_at", since);
  if (countErr) return json({ error: "Login service error." }, 500);
  if ((failures ?? 0) >= MAX_FAILURES) {
    return json({ error: `Too many wrong PINs. Try again in ${LOCKOUT_MINUTES} minutes.` }, 429);
  }

  // Same match the app always used: case-insensitive exact name + exact PIN, exactly one row.
  const escapedName = name.replace(/[\\%_]/g, (c) => "\\" + c);
  const { data: users, error: userErr } = await admin.from("users")
    .select("*")
    .ilike("name", escapedName)
    .eq("pin", pin)
    .limit(2);
  if (userErr) return json({ error: "Login service error." }, 500);

  if (!users || users.length !== 1) {
    await admin.from("login_attempts").insert({ name_key: nameKey });
    return json({ error: "Invalid Name or PIN." }, 401);
  }
  const user = users[0];
  await admin.from("login_attempts").delete().eq("name_key", nameKey);

  // One Supabase Auth user per Office OS user, keyed by the stable users.id (not the name).
  const email = `${user.id}@staff.officeos.local`;
  const { error: createErr } = await admin.auth.admin.createUser({
    email,
    email_confirm: true,
    app_metadata: { app_user_id: user.id },
  });
  if (createErr && createErr.code !== "email_exists") {
    console.error("createUser failed:", createErr);
    return json({ error: "Login service error." }, 500);
  }

  const { data: link, error: linkErr } = await admin.auth.admin.generateLink({ type: "magiclink", email });
  if (linkErr || !link?.properties?.hashed_token) {
    console.error("generateLink failed:", linkErr);
    return json({ error: "Login service error." }, 500);
  }
  if (link.user?.app_metadata?.app_user_id !== user.id) {
    await admin.auth.admin.updateUserById(link.user.id, {
      app_metadata: { ...link.user.app_metadata, app_user_id: user.id },
    });
  }

  return json({ user, token_hash: link.properties.hashed_token });
});
