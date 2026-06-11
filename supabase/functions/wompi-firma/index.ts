// wompi-firma — genera la firma de integridad del checkout de Wompi.
// Requiere usuario con sesión (Verify JWT: ON).
// Secret necesario: WOMPI_INTEGRITY_SECRET (Settings → Edge Functions → Secrets)

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, content-type, apikey",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  try {
    const { reference, amountInCents } = await req.json();
    const secret = Deno.env.get("WOMPI_INTEGRITY_SECRET") ?? "";
    if (!secret) {
      return new Response(JSON.stringify({ error: "Falta WOMPI_INTEGRITY_SECRET" }),
        { status: 500, headers: { ...CORS, "content-type": "application/json" } });
    }
    if (!reference || !amountInCents || !/^[0-9a-f-]{36}_\d+$/i.test(reference)) {
      return new Response(JSON.stringify({ error: "Datos inválidos" }),
        { status: 400, headers: { ...CORS, "content-type": "application/json" } });
    }
    const txt = `${reference}${amountInCents}COP${secret}`;
    const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(txt));
    const signature = [...new Uint8Array(buf)]
      .map((b) => b.toString(16).padStart(2, "0")).join("");
    return new Response(JSON.stringify({ signature }),
      { headers: { ...CORS, "content-type": "application/json" } });
  } catch (_e) {
    return new Response(JSON.stringify({ error: "Petición inválida" }),
      { status: 400, headers: { ...CORS, "content-type": "application/json" } });
  }
});
