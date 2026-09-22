import { supabase } from "./supabase";

// Calls the Anthropic Messages API through the `claude-proxy` Supabase Edge Function.
// The API key lives only on the server (Edge Function secret) and never reaches the browser.
// `body` has the same shape as a Messages API request; the response is returned unchanged
// (`{ content: [...] }`), so the calling code reads it exactly as before.
export async function callClaude(body) {
  const { data, error } = await supabase.functions.invoke("claude-proxy", { body });
  if (error) throw new Error("richiesta fallita");
  return data;
}
