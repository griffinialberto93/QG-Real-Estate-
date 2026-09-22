import React, { useEffect, useState } from "react";
import { supabase, supabaseConfigured } from "./lib/supabase";

// Login by email link (no passwords). Only users already created in Supabase Auth can sign in,
// and only emails listed in `team_members` can read or write data (enforced by RLS).

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
  const [email, setEmail] = useState("");
  const [stato, setStato] = useState("");

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
    const invia = async (e) => {
      e.preventDefault();
      setStato("Invio in corso…");
      const { error } = await supabase.auth.signInWithOtp({
        email: email.trim(),
        options: { shouldCreateUser: false, emailRedirectTo: window.location.origin },
      });
      setStato(error ? "Accesso non consentito per questa email." : "Controlla la posta: ti ho inviato il link di accesso.");
    };
    return (
      <div style={box}>
        <form style={card} onSubmit={invia}>
          <div style={{ fontSize: 17, fontWeight: 600 }}>QG Real Estate</div>
          <div style={{ fontSize: 12.5, color: "#8b91a0", marginTop: 4 }}>Accedi con la tua email aziendale.</div>
          <input style={input} type="email" required value={email} onChange={(e) => setEmail(e.target.value)} placeholder="nome@qgrealestate.it" />
          <button style={btn} type="submit">Invia link di accesso</button>
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
