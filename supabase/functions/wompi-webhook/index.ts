// wompi-webhook — recibe los eventos de Wompi y activa/renueva la licencia.
// URL para Wompi (Programadores → URL de eventos):
//   https://xdydorreyvkenbifefus.supabase.co/functions/v1/wompi-webhook
// IMPORTANTE: desplegar con Verify JWT: OFF (Wompi no manda token de Supabase).
// Secret necesario: WOMPI_EVENTS_SECRET (el "Secreto de eventos" de Wompi).

import { createClient } from "npm:@supabase/supabase-js@2";

// Precios en centavos COP por plan y periodo (1, 6 o 12 meses).
// Deben coincidir con PRECIOS de docs/cuenta.html.
const PRECIOS: Record<string, Record<string, number>> = {
  imagine:  { "1": 250_000_00, "6": 1_275_000_00, "12": 2_250_000_00 },
  guardian: { "1": 250_000_00, "6": 1_275_000_00, "12": 2_250_000_00 },
  completo: { "1": 450_000_00, "6": 2_295_000_00, "12": 4_050_000_00 },
};
const DIAS: Record<string, number> = { "1": 30, "6": 183, "12": 365 };

async function sha256hex(s: string): Promise<string> {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

Deno.serve(async (req) => {
  let event: any;
  try { event = await req.json(); } catch { return new Response("json inválido", { status: 400 }); }

  // ── Verificación de firma del evento (anti-falsificación) ────────────────
  const secret = Deno.env.get("WOMPI_EVENTS_SECRET") ?? "";
  const props: string[] = event?.signature?.properties ?? [];
  let chain = "";
  for (const p of props) {
    chain += String(p.split(".").reduce((o: any, k: string) => o?.[k], event?.data) ?? "");
  }
  chain += String(event?.timestamp ?? "") + secret;
  const checksum = await sha256hex(chain);
  if (!secret || checksum.toUpperCase() !== String(event?.signature?.checksum ?? "").toUpperCase()) {
    return new Response("firma inválida", { status: 401 });
  }

  // ── Solo pagos aprobados ──────────────────────────────────────────────────
  const tx = event?.data?.transaction;
  if (event?.event !== "transaction.updated" || tx?.status !== "APPROVED") {
    return new Response("ignorado");
  }

  // Referencia: uid_plan_meses_timestamp (ej: "a1b2..._completo_6_171234")
  const partes = String(tx.reference ?? "").split("_");
  const userId = partes[0];
  const plan   = partes[1] ?? "";
  const meses  = partes[2] ?? "";
  if (!/^[0-9a-f-]{36}$/i.test(userId)) return new Response("referencia inválida", { status: 400 });
  if (!PRECIOS[plan] || !PRECIOS[plan][meses]) return new Response("plan inválido", { status: 400 });

  // El monto pagado debe coincidir con el precio del plan (anti-manipulación).
  if (Number(tx.amount_in_cents) !== PRECIOS[plan][meses] || tx.currency !== "COP") {
    return new Response("monto no coincide con el plan", { status: 400 });
  }

  const sb = createClient(
    Deno.env.get("SUPABASE_URL")!,
    Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
  );

  // Renueva desde el vencimiento si la licencia sigue viva; si no, desde hoy.
  const { data: lic } = await sb.from("licenses")
    .select("expires_at").eq("user_id", userId).maybeSingle();
  const base = lic && new Date(lic.expires_at) > new Date()
    ? new Date(lic.expires_at) : new Date();
  base.setDate(base.getDate() + DIAS[meses]);

  const { error } = await sb.from("licenses").upsert({
    user_id: userId,
    status: "activa",
    plan: plan,                       // imagine | guardian | completo
    expires_at: base.toISOString(),
    updated_at: new Date().toISOString(),
  });
  if (error) return new Response("error bd: " + error.message, { status: 500 });

  return new Response("licencia activada");
});
