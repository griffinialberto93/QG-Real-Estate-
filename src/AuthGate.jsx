import React, { useEffect, useState } from "react";
import { supabase, supabaseConfigured } from "./lib/supabase";

// Login con nome utente e password: non viene inviata nessuna email. Gli account li crea un
// amministratore dalla dashboard Supabase; solo le email presenti in `team_members` possono
// leggere o scrivere i dati (policy RLS).

// Chi scrive solo il nome utente ("alberto") accede all'account "alberto@qgrealestate.it".
const DOMINIO_PREDEFINITO = "@qgrealestate.it";

const box = {
  minHeight: "100%", display: "flex", alignItems: "center", justifyContent: "center",
  background: "#12151a", color: "#e9e7df", fontFamily: "Inter, -apple-system, sans-serif", padding: 16,
};
const card = {
  width: "100%", maxWidth: 380, background: "#191d24", border: "1px solid #2a3040",
  borderRadius: 10, padding: 24,
};
const input = {
  width: "100%", boxSizing: "border-box", background: "#1f2430", border: "1px solid #2a3040",
  color: "#e9e7df", borderRadius: 6, padding: "9px 10px", fontSize: 14, margin: "10px 0",
};
const btn = {
  width: "100%", background: "#c19a5b", color: "#1a1509", border: "none", borderRadius: 6,
  padding: "9px 14px", fontSize: 14, fontWeight: 600, cursor: "pointer",
};

export default function AuthGate({ children }) {
  const [session, setSession] = useState(undefined);
  const [utente, setUtente] = useState("");
  const [password, setPassword] = useState("");
  const [stato, setStato] = useState("");
  const [inCorso, setInCorso] = useState(false);

  useEffect(() => {
    if (!supabaseConfigured) return;
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_e, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  if (!supabaseConfigured) {
    return (
      <div style={box}>
        <div style={card}>
          <b>Configurazione mancante</b>
          <p style={{ fontSize: 13, color: "#8b91a0" }}>
            Crea il file <code>.env.local</code> partendo da <code>.env.example</code> con URL e chiave anon del progetto Supabase, poi riavvia <code>npm run dev</code>.
          </p>
        </div>
      </div>
    );
  }

  if (session === undefined) return <div style={box}>Caricamento…</div>;

  if (!session) {
    const entra = async (e) => {
      e.preventDefault();
      const nome = utente.trim().toLowerCase();
      const emailAccount = nome.includes("@") ? nome : nome + DOMINIO_PREDEFINITO;
      setInCorso(true);
      setStato("");
      const { error } = await supabase.auth.signInWithPassword({ email: emailAccount, password });
      setInCorso(false);
      if (error) setStato("Nome utente o password non validi.");
    };
    return (
      <div style={box}>
        <form style={card} onSubmit={entra}>
          <div style={{ fontSize: 17, fontWeight: 600 }}>QG Real Estate</div>
          <div style={{ fontSize: 12.5, color: "#8b91a0", marginTop: 4 }}>Accedi con nome utente e password.</div>
          <input style={input} type="text" required autoFocus autoComplete="username" value={utente} onChange={(e) => setUtente(e.target.value)} placeholder="nome utente" />
          <input style={input} type="password" required autoComplete="current-password" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="password" />
          <button style={btn} type="submit" disabled={inCorso}>{inCorso ? "Accesso in corso…" : "Entra"}</button>
          {stato && <div style={{ fontSize: 12.5, color: "#8b91a0", marginTop: 12 }}>{stato}</div>}
        </form>
      </div>
    );
  }

  return (
    <>
      {children}
      <button
        onClick={() => supabase.auth.signOut()}
        title={session.user.email}
        style={{
          position: "fixed", right: 12, bottom: 12, fontSize: 11, padding: "5px 10px",
          background: "#1f2430", color: "#8b91a0", border: "1px solid #2a3040", borderRadius: 6, cursor: "pointer",
        }}
      >
        Esci
      </button>
    </>
  );
}
