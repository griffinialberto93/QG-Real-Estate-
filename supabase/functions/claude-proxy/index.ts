// Supabase Edge Function: proxy verso l'Anthropic Messages API.
//
// Perché esiste: nell'artifact le chiamate a Claude funzionavano senza chiave; fuori dall'artifact
// serve una API key, che NON deve mai finire nel browser. La chiave sta qui come secret
// (ANTHROPIC_API_KEY) e la funzione accetta richieste solo da utenti del team.
//
// Deploy:
//   supabase secrets set ANTHROPIC_API_KEY=sk-ant-...
//   supabase secrets set CLAUDE_MODEL=claude-sonnet-5   (opzionale)
//   supabase functions deploy claude-proxy

import { createClient } from "npm:@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const MAX_TOKENS_CAP = 2000;
// Solo lo strumento di ricerca web usato dall'app (quotazioni titoli e cambio USD/EUR).
const ALLOWED_TOOLS = new Set(["web_search"]);

const json = (status: number, body: unknown) =>
  new Response(JSON.stringify(body), { status, headers: { ...corsHeaders, "Content-Type": "application/json" } });

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  if (req.method !== "POST") return json(405, { error: "method not allowed" });

  // 1) L'utente deve essere autenticato e nel team.
  const authHeader = req.headers.get("Authorization") ?? "";
  const supabase = createClient(
    Deno.env.get("SUPABASE_URL")!,
    Deno.env.get("SUPABASE_ANON_KEY")!,
    { global: { headers: { Authorization: authHeader } } },
  );
  const { data: isMember, error: memberError } = await supabase.rpc("is_team_member");
  if (memberError || isMember !== true) return json(403, { error: "non autorizzato" });

  // 2) Richiesta ripulita: modello e limiti li decide il server, non il client.
  let input: Record<string, unknown>;
  try {
    input = await req.json();
  } catch {
    return json(400, { error: "body non valido" });
  }
  const messages = input.messages;
  if (!Array.isArray(messages) || messages.length === 0) return json(400, { error: "messages mancante" });

  const tools = Array.isArray(input.tools)
    ? (input.tools as Array<{ name?: string }>).filter((t) => t && ALLOWED_TOOLS.has(String(t.name)))
    : undefined;

  const body: Record<string, unknown> = {
    model: Deno.env.get("CLAUDE_MODEL") ?? "claude-sonnet-5",
    max_tokens: Math.min(Number(input.max_tokens) || 1000, MAX_TOKENS_CAP),
    messages,
  };
  if (typeof input.system === "string") body.system = input.system;
  if (tools && tools.length) body.tools = tools;

  // 3) Inoltro ad Anthropic.
  const apiKey = Deno.env.get("ANTHROPIC_API_KEY");
  if (!apiKey) return json(500, { error: "ANTHROPIC_API_KEY non configurata" });

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify(body),
  });

  const data = await res.json().catch(() => ({ error: "risposta non leggibile" }));
  return json(res.ok ? 200 : res.status, data);
});
