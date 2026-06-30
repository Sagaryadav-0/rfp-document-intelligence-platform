import { signToken } from "../shared.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const { password, type } = await req.json();
    const appPassword = Deno.env.get("APP_PASSWORD");
const adminPassword = Deno.env.get("ADMIN_PASSWORD");

if (!appPassword || !adminPassword) {
  throw new Error("Required environment variables are not configured.");
}

    if (type === "admin") {
      const ok = typeof password === "string" && password === adminPassword;
      if (!ok) {
        return new Response(JSON.stringify({ ok: false }), {
          status: 401,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        });
      }
      const token = await signToken({ role: "admin", exp: Date.now() + 15 * 60 * 1000 });
      return new Response(JSON.stringify({ ok: true, token }), {
        status: 200,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const ok = typeof password === "string" && password === appPassword;
    if (!ok) {
      return new Response(JSON.stringify({ ok: false }), {
        status: 401,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const token = await signToken({ role: "app", exp: Date.now() + 15 * 60 * 1000 });
    return new Response(JSON.stringify({ ok: true, token }), {
      status: 200,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (err) {
  console.error(err);

  return new Response(
    JSON.stringify({
      ok: false,
      error: "Internal server error",
    }),
    {
      status: 500,
      headers: {
        ...corsHeaders,
        "Content-Type": "application/json",
      },
    }
  );
}
});
