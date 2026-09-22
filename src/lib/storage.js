import { supabase } from "./supabase";

// Drop-in replacement for the artifact's `window.storage` (get/set of a JSON string by key),
// backed by the `app_state` table on Supabase.
//
// The whole workspace (liquidità, operazioni, immobili, bilancio) is stored as one JSON document,
// exactly as it was in the artifact, so the export/import backup format stays unchanged.
//
// Concurrency: several partners can have the app open at once. Every row carries a `version`;
// a save only succeeds if nobody else saved in between (optimistic locking via the
// `save_app_state` SQL function). Otherwise it throws a ConflictError instead of silently
// overwriting a colleague's changes.

export class ConflictError extends Error {
  constructor() {
    super("I dati sono stati modificati da un altro utente.");
    this.name = "ConflictError";
  }
}

const versions = new Map(); // key -> last version seen by this browser tab

export const storage = {
  async get(key) {
    const { data, error } = await supabase
      .from("app_state")
      .select("data, version")
      .eq("key", key)
      .maybeSingle();
    if (error) throw error;
    if (!data) {
      versions.set(key, 0);
      return null;
    }
    versions.set(key, data.version);
    return { key, value: JSON.stringify(data.data) };
  },

  async set(key, value) {
    const { data, error } = await supabase.rpc("save_app_state", {
      p_key: key,
      p_data: JSON.parse(value),
      p_expected_version: versions.get(key) ?? 0,
    });
    if (error) {
      if (error.message && error.message.includes("CONFLICT")) throw new ConflictError();
      throw error;
    }
    versions.set(key, data);
    return { key, value };
  },
};
