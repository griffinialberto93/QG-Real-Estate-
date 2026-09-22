import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { storage, ConflictError } from "./lib/storage";
import { callClaude } from "./lib/claude";
import {
  AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, ReferenceLine, BarChart, Bar, Cell,
} from "recharts";
import {
  Building2, TrendingUp, TrendingDown, Plus, Trash2, ChevronLeft, ChevronRight,
  AlertTriangle, CheckCircle2, Landmark, Users, Wallet, LayoutDashboard,
  ListTree, Settings2, Save, Clock, Receipt,
  PiggyBank, ShieldAlert, CircleDollarSign, CalendarClock, Upload, FileUp,
  KeyRound, Home, ArrowDownLeft, ArrowUpRight, Pencil,
  LineChart as LineChartIcon, Scale, Sparkles, SendHorizontal, RefreshCw,
  Download,
} from "lucide-react";

/* ============================================================
   TOKENS / STYLE
   ============================================================ */
const GlobalStyle = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

    .qg-root {
      --bg: #12151a;
      --panel: #191d24;
      --panel-2: #1f2430;
      --panel-3: #262c3a;
      --border: #2a3040;
      --border-soft: #232836;
      --text: #e9e7df;
      --text-muted: #8b91a0;
      --text-dim: #5c6274;
      --accent: #c19a5b;
      --accent-soft: rgba(193, 154, 91, 0.14);
      --accent-strong: #d9b06f;
      --positive: #5fac8d;
      --positive-soft: rgba(95, 172, 141, 0.14);
      --negative: #c1594a;
      --negative-soft: rgba(193, 89, 74, 0.14);
      --warning: #d4a857;
      --warning-soft: rgba(212, 168, 87, 0.14);
      font-family: 'Inter', -apple-system, sans-serif;
      background: var(--bg);
      color: var(--text);
      min-height: 100%;
      width: 100%;
      -webkit-font-smoothing: antialiased;
    }
    .qg-root * { box-sizing: border-box; }
    .qg-display { font-family: 'Fraunces', Georgia, serif; }
    .qg-mono { font-family: 'IBM Plex Mono', ui-monospace, monospace; }
    .qg-scroll::-webkit-scrollbar { width: 8px; height: 8px; }
    .qg-scroll::-webkit-scrollbar-thumb { background: var(--border); border-radius: 4px; }
    .qg-scroll::-webkit-scrollbar-track { background: transparent; }

    .qg-input {
      background: var(--panel-2);
      border: 1px solid var(--border);
      color: var(--text);
      border-radius: 6px;
      padding: 8px 10px;
      font-size: 13px;
      font-family: 'IBM Plex Mono', monospace;
      width: 100%;
      outline: none;
      transition: border-color .15s;
    }
    .qg-input:focus { border-color: var(--accent); }
    .qg-input::placeholder { color: var(--text-dim); }
    select.qg-input { font-family: 'Inter', sans-serif; }

    .qg-label {
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      color: var(--text-muted);
      margin-bottom: 5px;
      display: block;
      font-weight: 500;
    }

    .qg-btn {
      display: inline-flex; align-items: center; gap: 6px;
      background: var(--panel-3); color: var(--text);
      border: 1px solid var(--border); border-radius: 6px;
      padding: 8px 14px; font-size: 13px; font-weight: 500;
      cursor: pointer; transition: all .15s; font-family: 'Inter', sans-serif;
    }
    .qg-btn:hover { border-color: var(--accent); color: var(--accent-strong); }
    .qg-btn-primary {
      background: var(--accent); color: #1a1509; border: 1px solid var(--accent);
    }
    .qg-btn-primary:hover { background: var(--accent-strong); color: #1a1509; }
    .qg-btn-ghost { background: transparent; border: 1px solid transparent; }
    .qg-btn-ghost:hover { background: var(--panel-2); }
    .qg-btn-danger { color: var(--negative); }
    .qg-btn-danger:hover { border-color: var(--negative); color: var(--negative); }
    .qg-btn:disabled { opacity: 0.4; cursor: not-allowed; }
    .qg-btn:disabled:hover { border-color: var(--border); color: var(--text); }

    .qg-panel {
      background: var(--panel);
      border: 1px solid var(--border-soft);
      border-radius: 10px;
    }

    .qg-badge {
      display: inline-flex; align-items: center; gap: 5px;
      padding: 3px 9px; border-radius: 20px; font-size: 11px;
      font-weight: 600; letter-spacing: 0.02em; white-space: nowrap;
    }
  `}</style>
);

/* ============================================================
   HELPERS
   ============================================================ */
const uid = () => Math.random().toString(36).slice(2, 10);

const ymToDate = (ym) => { const [y, m] = ym.split("-").map(Number); return new Date(y, m - 1, 1); };
const dateToYm = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
const addMonths = (ym, n) => { const d = ymToDate(ym); d.setMonth(d.getMonth() + n); return dateToYm(d); };
const monthDiff = (a, b) => { const da = ymToDate(a), db = ymToDate(b); return (db.getFullYear() - da.getFullYear()) * 12 + (db.getMonth() - da.getMonth()); };
const currentYm = () => dateToYm(new Date());
const monthLabel = (ym) => { if (!ym) return "—"; const d = ymToDate(ym); return d.toLocaleDateString("it-IT", { month: "short", year: "2-digit" }).replace(".", ""); };
const monthLabelLong = (ym) => { if (!ym) return "—"; const d = ymToDate(ym); const s = d.toLocaleDateString("it-IT", { month: "long", year: "numeric" }); return s.charAt(0).toUpperCase() + s.slice(1); };

// Punctual (day-level) dates for one-off expenses/movements, stored as "YYYY-MM-DD".
// The monthly projection engine still buckets these by month via ymdToYm().
const todayStr = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; };
const ymdToYm = (ymd) => (ymd ? ymd.slice(0, 7) : ymd);
// Migrates legacy "YYYY-MM" values (or ym-style helper output) into a full "YYYY-MM-DD" date, defaulting to the 1st.
const toDateStr = (v) => { if (!v) return v; return v.length === 7 ? `${v}-01` : v; };
const dateLabel = (ymd) => {
  if (!ymd) return "—";
  const [y, m, d] = ymd.split("-").map(Number);
  const dt = new Date(y, m - 1, d);
  return dt.toLocaleDateString("it-IT", { day: "numeric", month: "short", year: "numeric" }).replace(".", "");
};
// Whole days from date a to date b (both "YYYY-MM-DD"), positive if b is after a.
const monthDiffDays = (a, b) => {
  if (!a || !b) return 0;
  const [ya, ma, da] = a.split("-").map(Number);
  const [yb, mb, db] = b.split("-").map(Number);
  const ta = Date.UTC(ya, ma - 1, da);
  const tb = Date.UTC(yb, mb - 1, db);
  return Math.round((tb - ta) / 86400000);
};
const daysInMonth = (ym) => { const [y, m] = ym.split("-").map(Number); return new Date(y, m, 0).getDate(); };
// Actual charge date of a recurring cost inside a given month. An empty/invalid day means
// "end of month", which keeps the conservative behaviour (the cost is assumed still to come).
const dueDateForMonth = (ym, giorno) => {
  const dim = daysInMonth(ym);
  const g = Math.floor(Number(giorno));
  const d = Number.isFinite(g) && g >= 1 ? Math.min(g, dim) : dim;
  return `${ym}-${String(d).padStart(2, "0")}`;
};

const fmtEUR = (n, opts = {}) => {
  const v = Number(n) || 0;
  return new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR", maximumFractionDigits: opts.decimals ? 0 : 0, minimumFractionDigits: 0 }).format(v);
};
const fmtEURSigned = (n) => {
  const v = Number(n) || 0;
  const s = fmtEUR(Math.abs(v));
  return v >= 0 ? `+${s}` : `-${s}`;
};
const fmtCompact = (n) => {
  const v = Number(n) || 0;
  const abs = Math.abs(v);
  if (abs >= 1000000) return (v / 1000000).toLocaleString("it-IT", { maximumFractionDigits: 2 }) + "M €";
  if (abs >= 1000) return (v / 1000).toLocaleString("it-IT", { maximumFractionDigits: 0 }) + "k €";
  return fmtEUR(v);
};
const fmtPct = (n, decimals = 1) => {
  const v = Number(n);
  if (!Number.isFinite(v)) return "—";
  return v.toLocaleString("it-IT", { minimumFractionDigits: decimals, maximumFractionDigits: decimals }) + "%";
};

/* ============================================================
   DOMAIN CONSTANTS
   ============================================================ */
const STATI = [
  { key: "pipeline", label: "Pipeline", color: "var(--warning)" },
  { key: "in_corso", label: "In corso", color: "var(--accent)" },
  { key: "completata", label: "Completata", color: "var(--positive)" },
  { key: "abbandonata", label: "Abbandonata", color: "var(--text-dim)" },
];
const statoInfo = (key) => STATI.find((s) => s.key === key) || STATI[0];

const CAT_ONEOFF = ["Acquisto", "Ristrutturazione / Lavori", "Notaio e Imposte", "Commissioni Agenzia", "Progettazione / Tecnici", "Altro"];
const CAT_RICORRENTE = ["Manutenzione", "Utenze", "Assicurazione", "Gestione / Amministrazione", "Altro"];

const CAT_FISSI_UNA_TANTUM = ["Attrezzature e arredi", "Consulenze straordinarie", "Imposte e tasse societarie", "Altro"];
const CAT_FISSI_RICORRENTE = ["Stipendi e collaboratori", "Affitto e utenze sede", "Consulenze fisse", "Software e abbonamenti", "Assicurazioni aziendali", "Altro"];

const CAT_RENTAL_ONEOFF = ["Acquisto", "Ristrutturazione / Lavori", "Notaio e Imposte", "Arredo e allestimento", "Commissioni Agenzia", "Altro"];
const CAT_RENTAL_RICORRENTE = ["Gestione / Property management", "Pulizie e biancheria", "Utenze", "IMU", "Assicurazione", "Manutenzione", "Condominio", "Commissioni piattaforme", "Altro"];

const emptyOperation = () => ({
  id: uid(),
  nome: "",
  indirizzo: "",
  stato: "pipeline",
  destinazione: "flipping", // "flipping" (rivendita) | "reddito" (tenuto e affittato)
  frazionamento: false,     // se true si vende in più unità, con date separate
  meseInizio: currentYm(),
  meseVenditaPrevista: addMonths(currentYm(), 9),
  prezzoVenditaPrevisto: "",
  vendite: [],              // [{ id, descrizione, importo, mese }] usato quando frazionamento = true
  costiUnaTantum: [],
  costiRicorrenti: [],
  copertura: {
    liquiditaQG: "",
    // Per il flipping è un prestito ponte (soli interessi, rimborso alla vendita).
    // Per le operazioni a reddito è un mutuo che ammortizza, con eventuale pre-ammortamento.
    prestito: { attivo: false, importo: "", tassoAnnuo: "5", meseErogazione: currentYm(), rimborsoAScadenza: true, durataMesi: "240", mesiPreammortamento: "" },
    investitori: { attivo: false, lista: [] },
  },
  note: "",
});

const emptyLiquidity = () => ({
  conti: [{ id: uid(), nome: "Conto principale", saldo: "", vincolato: "" }],
  dataRiferimento: todayStr(),
  soglioMinimo: "0",
  movimenti: [],
  partite: [], // debiti/crediti con data singola di incasso/rimborso
  portafoglio: { titoli: [] }, // posizioni azionarie: riserva liquidabile ma non cassa operativa
  fx: { usdToEur: "", aggiornato: "" }, // tasso USD→EUR, recuperato o inserito a mano
  aliquoteFiscali: { ires: "24", irap: "3.9", sostitutiva: "26" },
  costiFissi: { unaTantum: [], ricorrenti: [] },
});

// A buy-to-let / short-let property kept on the books and rented out (not flipped).
const emptyRental = () => ({
  id: uid(),
  nome: "",
  indirizzo: "",
  operazioneId: "", // se valorizzato, costi d'acquisto e mutuo arrivano da quell'operazione
  tipoAffitto: "breve", // "breve" (affitto breve turistico) | "lungo" (canone mensile)
  attivo: true,
  meseInizioLocazione: currentYm(),
  costiUnaTantum: [],   // { id, categoria, importo, data } — acquisto, lavori, arredo…
  costiRicorrenti: [],  // { id, categoria, importoMensile, meseInizio, meseFine }
  ricavi: {
    // Affitto breve: tariffa media a notte × tasso di occupazione.
    tariffaNotte: "",
    occupazionePct: "",
    // Affitto lungo (o ricavo diretto): canone mensile.
    canoneMensile: "",
  },
  mutuo: { attivo: false, importo: "", tassoAnnuo: "", durataMesi: "", mesiPreammortamento: "", meseErogazione: currentYm() },
  note: "",
});

// Fills in any field missing from older saved data so the UI never renders on undefined values.
// Also migrates the older single-amount investor shape ({ importo, quotaPct }) into a list of
// named investors, so previously saved operations keep working after this change.
function normalizeOperation(op) {
  const base = emptyOperation();
  const oldInvestitori = (op.copertura && op.copertura.investitori) || {};
  let lista = oldInvestitori.lista;
  if (!lista && Number(oldInvestitori.importo) > 0) {
    lista = [{ id: uid(), nome: "Investitore 1", importo: oldInvestitori.importo }];
  }
  return {
    ...base,
    ...op,
    destinazione: op.destinazione === "reddito" ? "reddito" : "flipping",
    frazionamento: !!op.frazionamento,
    vendite: (op.vendite || []).map((v) => ({
      id: v.id || uid(),
      descrizione: v.descrizione || "",
      importo: v.importo ?? "",
      mese: v.mese || op.meseVenditaPrevista || currentYm(),
    })),
    costiUnaTantum: (op.costiUnaTantum || []).map((c) => {
      const { mese, ...rest } = c;
      return { ...rest, data: c.data || toDateStr(mese) || todayStr() };
    }),
    costiRicorrenti: (op.costiRicorrenti || []).map((c) => ({ ...c, giornoAddebito: c.giornoAddebito ?? "" })),
    copertura: {
      ...base.copertura,
      ...(op.copertura || {}),
      prestito: { ...base.copertura.prestito, ...((op.copertura && op.copertura.prestito) || {}) },
      investitori: {
        attivo: !!oldInvestitori.attivo,
        // meseVersamento: quando il capitale dell'investitore arriva sul conto QG.
        lista: (lista || []).map((i) => ({ ...i, meseVersamento: i.meseVersamento || op.meseInizio || currentYm() })),
      },
    },
  };
}

// Sale events of an operation, normalised to a single list: one entry for a plain resale,
// several (with their own dates) when the property is split into units ("frazionamento").
function venditeOperazione(op) {
  if (op.destinazione === "reddito") return [];
  if (op.frazionamento) {
    return (op.vendite || [])
      .filter((v) => v.mese && Number(v.importo) > 0)
      .map((v) => ({ id: v.id, descrizione: v.descrizione || "Unità", importo: Number(v.importo) || 0, mese: v.mese }))
      .sort((a, b) => (a.mese < b.mese ? -1 : 1));
  }
  if (!op.meseVenditaPrevista || !(Number(op.prezzoVenditaPrevisto) > 0)) return [];
  return [{ id: "unica", descrizione: "Vendita", importo: Number(op.prezzoVenditaPrevisto) || 0, mese: op.meseVenditaPrevista }];
}

// Total expected revenue and the month of the last sale (when the operation actually closes).
const ricavoTotaleOperazione = (op) => venditeOperazione(op).reduce((s, v) => s + v.importo, 0);
const meseChiusuraOperazione = (op) => {
  const v = venditeOperazione(op);
  return v.length ? v[v.length - 1].mese : (op.meseVenditaPrevista || op.meseInizio);
};

// Fills in any field missing from older saved liquidity data, migrating legacy month-only
// dates on one-off fixed costs and extra movements into full "YYYY-MM-DD" dates.
function normalizeLiquidity(liq) {
  const base = emptyLiquidity();
  const src = liq || {};
  return {
    ...base,
    ...src,
    aliquoteFiscali: { ...base.aliquoteFiscali, ...(src.aliquoteFiscali || {}) },
    costiFissi: {
      unaTantum: (src.costiFissi?.unaTantum || []).map((c) => {
        const { mese, ...rest } = c;
        return { ...rest, data: c.data || toDateStr(mese) || todayStr() };
      }),
      ricorrenti: (src.costiFissi?.ricorrenti || []).map((c) => ({ ...c, giornoAddebito: c.giornoAddebito ?? "" })),
    },
    movimenti: (src.movimenti || []).map((m) => ({ ...m, data: toDateStr(m.data) || todayStr() })),
    partite: (src.partite || []).map((p) => ({
      id: p.id || uid(),
      tipo: p.tipo === "debito" ? "debito" : "credito",
      controparte: p.controparte || "",
      importo: p.importo ?? "",
      data: toDateStr(p.data) || todayStr(),
      saldato: !!p.saldato,
      descrizione: p.descrizione || "",
    })),
    portafoglio: {
      titoli: (src.portafoglio?.titoli || []).map((t) => ({
        id: t.id || uid(),
        nome: t.nome || "",
        ticker: t.ticker || "",
        quantita: t.quantita ?? "",
        prezzoCarico: t.prezzoCarico ?? "",
        prezzoAttuale: t.prezzoAttuale ?? "",
        valuta: t.valuta === "USD" ? "USD" : "EUR",
      })),
    },
    fx: { usdToEur: src.fx?.usdToEur ?? "", aggiornato: src.fx?.aggiornato ?? "" },
    // Data di riconciliazione: tutto ciò che è datato fino a qui è già dentro i saldi dei conti.
    dataRiferimento: toDateStr(src.dataRiferimento) || todayStr(),
    // Conti correnti: migra il vecchio saldo unico in un conto singolo.
    conti: Array.isArray(src.conti) && src.conti.length
      ? src.conti.map((c) => ({ id: c.id || uid(), nome: c.nome || "Conto", saldo: c.saldo ?? "", vincolato: c.vincolato ?? "" }))
      : [{ id: uid(), nome: "Conto principale", saldo: src.saldoAttuale ?? "", vincolato: "" }],
  };
}

// Gross balance across all accounts (what the bank shows).
const totaleLiquidita = (liq) => {
  if (Array.isArray(liq?.conti) && liq.conti.length) return liq.conti.reduce((s, c) => s + (Number(c.saldo) || 0), 0);
  return Number(liq?.saldoAttuale) || 0;
};
// Money sitting in the accounts that belongs to third parties (e.g. the owners' share of
// bookings collected by Smart Rent): it is in the bank, but it is not QG's to spend.
const totaleVincolato = (liq) => (liq?.conti || []).reduce((s, c) => s + (Number(c.vincolato) || 0), 0);
// Spendable cash: this is what drives the projection and the safety threshold.
const totaleDisponibile = (liq) => totaleLiquidita(liq) - totaleVincolato(liq);

// Simple two-branch management P&L, entered manually month by month.
const MESI_BREVI = ["Gen", "Feb", "Mar", "Apr", "Mag", "Giu", "Lug", "Ago", "Set", "Ott", "Nov", "Dic"];
const RAMI = [
  { key: "operazioni", label: "Lato Operazioni", color: "var(--accent)" },
  { key: "smartRent", label: "Smart Rent Solution", color: "var(--positive)" },
];
const mkKey = (i) => String(i + 1).padStart(2, "0");
const emptyMeseBilancio = () => ({ operazioni: { entrate: "", uscite: "", restituzione: "" }, smartRent: { entrate: "", uscite: "" } });

const emptyBilancio = () => ({ anni: {} });

function normalizeBilancio(b) {
  const src = b || {};
  const anni = {};
  Object.entries(src.anni || {}).forEach(([anno, mesi]) => {
    anni[anno] = {};
    for (let i = 0; i < 12; i++) {
      const k = mkKey(i);
      const m = (mesi && mesi[k]) || {};
      anni[anno][k] = {
        operazioni: { entrate: m.operazioni?.entrate ?? "", uscite: m.operazioni?.uscite ?? "", restituzione: m.operazioni?.restituzione ?? "" },
        smartRent: { entrate: m.smartRent?.entrate ?? "", uscite: m.smartRent?.uscite ?? "" },
      };
    }
  });
  return { anni };
}

// Aggregates a year of the branch P&L into per-month, per-branch and yearly totals.
// For the Operations branch, "restituzione capitale" is third-party capital repaid: it is
// subtracted from PROFIT only (branch and total) — it never reduces revenue (fatturato).
function bilancioAnnuale(bilancio, anno) {
  const mesi = (bilancio?.anni && bilancio.anni[anno]) || {};
  const perMese = [];
  const tot = {
    operazioni: { entrate: 0, uscite: 0, restituzione: 0 },
    smartRent: { entrate: 0, uscite: 0 },
  };
  for (let i = 0; i < 12; i++) {
    const k = mkKey(i);
    const cell = mesi[k] || emptyMeseBilancio();
    const opE = Number(cell.operazioni?.entrate) || 0;
    const opU = Number(cell.operazioni?.uscite) || 0;
    const opR = Number(cell.operazioni?.restituzione) || 0;
    const srE = Number(cell.smartRent?.entrate) || 0;
    const srU = Number(cell.smartRent?.uscite) || 0;
    tot.operazioni.entrate += opE; tot.operazioni.uscite += opU; tot.operazioni.restituzione += opR;
    tot.smartRent.entrate += srE; tot.smartRent.uscite += srU;
    const fatturato = opE + srE;           // fatturato lordo, sempre ≥ 0
    const uscite = opU + srU;
    const utile = fatturato - uscite - opR; // la restituzione capitale abbatte solo l'utile
    perMese.push({
      idx: i, mese: MESI_BREVI[i],
      opE, opU, opR, srE, srU,
      opUtile: opE - opU - opR, srUtile: srE - srU,
      fatturato, uscite, utile,
    });
  }
  const fatturato = tot.operazioni.entrate + tot.smartRent.entrate;
  const uscite = tot.operazioni.uscite + tot.smartRent.uscite;
  const utile = fatturato - uscite - tot.operazioni.restituzione;
  return {
    perMese, tot,
    fatturato, uscite, utile,
    margine: fatturato > 0 ? (utile / fatturato) * 100 : 0,
    opUtile: tot.operazioni.entrate - tot.operazioni.uscite - tot.operazioni.restituzione,
    srUtile: tot.smartRent.entrate - tot.smartRent.uscite,
  };
}

// Fills in any field missing from older saved rental data.
function normalizeRental(r) {
  const base = emptyRental();
  const src = r || {};
  return {
    ...base,
    ...src,
    operazioneId: src.operazioneId || "",
    costiUnaTantum: (src.costiUnaTantum || []).map((c) => ({ id: c.id || uid(), categoria: c.categoria || CAT_RENTAL_ONEOFF[0], importo: c.importo ?? "", data: c.data || todayStr() })),
    costiRicorrenti: (src.costiRicorrenti || []).map((c) => ({ id: c.id || uid(), categoria: c.categoria || CAT_RENTAL_RICORRENTE[0], importoMensile: c.importoMensile ?? "", meseInizio: c.meseInizio || src.meseInizioLocazione || currentYm(), meseFine: c.meseFine || "", giornoAddebito: c.giornoAddebito ?? "" })),
    ricavi: { ...base.ricavi, ...(src.ricavi || {}) },
    mutuo: { ...base.mutuo, ...(src.mutuo || {}) },
  };
}

/* ============================================================
   CALCULATION ENGINE
   ============================================================ */
// Flow maps are keyed by month and keep gross inflows and outflows apart, so the projection
// can show both the net balance and the entrate/uscite breakdown without losing information
// when a month contains both (e.g. a sale and the loan repayment that goes with it).
const addFlow = (map, ym, amt, voce) => {
  if (!ym || !amt) return;
  if (!map[ym]) map[ym] = { in: 0, out: 0, voci: [] };
  if (amt >= 0) map[ym].in += amt; else map[ym].out += -amt;
  // Ogni movimento porta con sé la sua etichetta, così la proiezione sa spiegare i propri numeri.
  map[ym].voci.push({ voce: voce || "Movimento", importo: amt });
};
const flowNet = (cell) => (cell ? cell.in - cell.out : 0);

// `cutoff` is the reconciliation date ("YYYY-MM-DD"): dated costs on or before it are already
// reflected in the account balances, so they must not be projected again.
function buildCostMap(op, cutoff) {
  const map = {};
  (op.costiUnaTantum || []).forEach((c) => {
    if (!c.data) return;
    if (cutoff && c.data <= cutoff) return;
    const ym = ymdToYm(c.data);
    map[ym] = map[ym] || { oneOff: 0, recurring: 0 };
    map[ym].oneOff += Number(c.importo) || 0;
  });
  const fineRicorrenti = meseChiusuraOperazione(op) || op.meseInizio;
  (op.costiRicorrenti || []).forEach((c) => {
    const start = c.meseInizio || op.meseInizio;
    const end = c.meseFine || fineRicorrenti;
    if (!start || !end) return;
    let cur = start, guard = 0;
    while (monthDiff(cur, end) >= 0 && guard < 400) {
      // Skip months whose charge date already falls inside the reconciled balances.
      if (!(cutoff && dueDateForMonth(cur, c.giornoAddebito) <= cutoff)) {
        map[cur] = map[cur] || { oneOff: 0, recurring: 0 };
        map[cur].recurring += Number(c.importoMensile) || 0;
      }
      cur = addMonths(cur, 1);
      guard++;
    }
  });
  return map;
}

function fundingShare(op) {
  const fabbisognoIniziale = (op.costiUnaTantum || []).reduce((s, c) => s + (Number(c.importo) || 0), 0);
  const liquiditaQG = Number(op.copertura?.liquiditaQG) || 0;
  const prestitoAttivo = !!op.copertura?.prestito?.attivo;
  const investitoriAttivo = !!op.copertura?.investitori?.attivo;
  const prestito = prestitoAttivo ? Number(op.copertura.prestito.importo) || 0 : 0;
  const listaInvestitori = investitoriAttivo ? (op.copertura?.investitori?.lista || []) : [];
  const investitori = listaInvestitori.reduce((s, i) => s + (Number(i.importo) || 0), 0);
  const coperturaTotale = liquiditaQG + prestito + investitori;
  const gap = fabbisognoIniziale - coperturaTotale;
  const pctQG = coperturaTotale > 0 ? liquiditaQG / coperturaTotale : (fabbisognoIniziale > 0 ? 1 : 0);
  return { fabbisognoIniziale, liquiditaQG, prestito, investitori, coperturaTotale, gap, pctQG };
}

// Full economics of an operation: costs, pre-tax margin, corporate taxes (IRES/IRAP),
// investor profit-share (proportional to their capital vs. total upfront cost), and the
// 26% substitute tax withheld on the investors' distribution.
function operationEconomics(op, liquidity) {
  const share = fundingShare(op);
  const aReddito = op.destinazione === "reddito";
  const ricavoTotale = ricavoTotaleOperazione(op);
  const meseChiusura = meseChiusuraOperazione(op);

  const totaleCostiRicorrenti = (op.costiRicorrenti || []).reduce((acc, c) => {
    const start = c.meseInizio || op.meseInizio, end = c.meseFine || meseChiusura || start;
    if (!start || !end) return acc;
    const mesi = Math.max(1, monthDiff(start, end) + 1);
    return acc + (Number(c.importoMensile) || 0) * mesi;
  }, 0);

  const prestito = op.copertura?.prestito?.attivo ? op.copertura.prestito : null;
  let interessiTotali = 0;
  if (prestito && Number(prestito.importo) > 0 && meseChiusura && !aReddito) {
    const mesi = Math.max(0, monthDiff(prestito.meseErogazione || op.meseInizio, meseChiusura));
    interessiTotali = (Number(prestito.importo) || 0) * (Number(prestito.tassoAnnuo) || 0) / 100 / 12 * mesi;
  }

  // Un'operazione a reddito non si chiude con una vendita: il suo rendimento si misura
  // nella scheda "Immobili a reddito", non come margine di compravendita.
  const marginLordo = aReddito ? 0 : ricavoTotale - share.fabbisognoIniziale - totaleCostiRicorrenti - interessiTotali;

  const iresPct = Number(liquidity?.aliquoteFiscali?.ires) || 0;
  const irapPct = Number(liquidity?.aliquoteFiscali?.irap) || 0;
  const sostitutivaPct = Number(liquidity?.aliquoteFiscali?.sostitutiva) || 0;

  const imposteSocietarie = marginLordo > 0 ? marginLordo * (iresPct + irapPct) / 100 : 0;
  const marginNetto = marginLordo - imposteSocietarie;

  const listaInvestitori = op.copertura?.investitori?.attivo ? (op.copertura.investitori.lista || []) : [];
  const investitoriDettaglio = listaInvestitori.map((inv) => {
    const importo = Number(inv.importo) || 0;
    const quotaPct = share.fabbisognoIniziale > 0 ? (importo / share.fabbisognoIniziale) * 100 : 0;
    const utileLordo = marginNetto > 0 ? marginNetto * (quotaPct / 100) : 0;
    const impostaSostitutiva = utileLordo > 0 ? utileLordo * sostitutivaPct / 100 : 0;
    const utileNetto = utileLordo - impostaSostitutiva;
    const totaleLiquidato = importo + utileNetto;
    return { id: inv.id, nome: inv.nome || "Investitore", importo, quotaPct, utileLordo, impostaSostitutiva, utileNetto, totaleLiquidato };
  });

  const quotaInvestitoriPct = investitoriDettaglio.reduce((s, i) => s + i.quotaPct, 0);
  const utileInvestitoriLordo = investitoriDettaglio.reduce((s, i) => s + i.utileLordo, 0);
  const impostaSostitutivaInvestitori = investitoriDettaglio.reduce((s, i) => s + i.impostaSostitutiva, 0);
  const utileInvestitoriNetto = investitoriDettaglio.reduce((s, i) => s + i.utileNetto, 0);
  const capitaleInvestitori = share.investitori;
  const totaleRestituitoInvestitori = capitaleInvestitori + utileInvestitoriNetto;
  const marginQG = marginNetto - utileInvestitoriLordo;

  return {
    share, totaleCostiRicorrenti, interessiTotali, marginLordo, iresPct, irapPct, sostitutivaPct,
    imposteSocietarie, marginNetto, quotaInvestitoriPct, utileInvestitoriLordo, investitoriDettaglio,
    impostaSostitutivaInvestitori, utileInvestitoriNetto, capitaleInvestitori, totaleRestituitoInvestitori, marginQG,
    aReddito, ricavoTotale, meseChiusura, vendite: venditeOperazione(op),
  };
}

function operationQGFlows(op, liquidity, opts = {}) {
  const { ignoreSale = false, horizonMonths = 72, cutoff = null } = opts;
  const flows = {};
  const aReddito = op.destinazione === "reddito";
  const nomeOp = op.nome || "Operazione senza nome";
  const costMap = buildCostMap(op, cutoff);
  const prestito = op.copertura?.prestito?.attivo ? op.copertura.prestito : null;
  const importoPrestito = prestito ? Number(prestito.importo) || 0 : 0;
  // A month is already settled when its last day falls on or before the reconciliation date.
  const giaRiconciliato = (ym) => cutoff && dueDateForMonth(ym, "") <= cutoff;

  // Uscite: i costi entrano per intero. Il capitale di soci e banca transita dal conto QG,
  // quindi viene registrato come entrata separata al momento in cui arriva: così si vede
  // anche lo scoperto temporaneo se il pagamento precede il versamento.
  Object.entries(costMap).forEach(([ym, vals]) => {
    if (vals.oneOff) addFlow(flows, ym, -vals.oneOff, `${nomeOp} · costi d'acquisto/lavori`);
    if (vals.recurring) addFlow(flows, ym, -vals.recurring, `${nomeOp} · costi ricorrenti`);
  });

  // Entrate: versamenti degli investitori.
  const listaInv = op.copertura?.investitori?.attivo ? (op.copertura.investitori.lista || []) : [];
  listaInv.forEach((inv) => {
    const amt = Number(inv.importo) || 0;
    const ym = inv.meseVersamento || op.meseInizio;
    if (amt && ym && !giaRiconciliato(ym)) addFlow(flows, ym, amt, `${nomeOp} · capitale versato da ${inv.nome || "investitore"}`);
  });

  // Entrata: erogazione del prestito / mutuo.
  if (importoPrestito > 0) {
    const ymEro = prestito.meseErogazione || op.meseInizio;
    if (ymEro && !giaRiconciliato(ymEro)) addFlow(flows, ymEro, importoPrestito, `${nomeOp} · erogazione ${aReddito ? "mutuo" : "prestito"}`);
  }

  const vendite = ignoreSale ? [] : venditeOperazione(op);
  const ricavoTot = vendite.reduce((s, v) => s + v.importo, 0);
  const econ = operationEconomics(op, liquidity);
  const meseChiusura = meseChiusuraOperazione(op);

  // Servizio del debito.
  if (importoPrestito > 0) {
    const start = prestito.meseErogazione || op.meseInizio;
    const tassoMensile = (Number(prestito.tassoAnnuo) || 0) / 100 / 12;
    if (aReddito) {
      // Mutuo che ammortizza: soli interessi durante il pre-ammortamento, poi rata piena.
      const durata = Math.max(0, Math.floor(Number(prestito.durataMesi) || 0));
      const mesiPre = Math.max(0, Math.min(Math.floor(Number(prestito.mesiPreammortamento) || 0), durata));
      const rataPre = importoPrestito * tassoMensile;
      const rataPiena = rataMutuo(importoPrestito, prestito.tassoAnnuo, durata - mesiPre);
      let cur = start, count = 0;
      while (count < durata && count < horizonMonths + 12) {
        if (!giaRiconciliato(cur)) addFlow(flows, cur, -(count < mesiPre ? rataPre : rataPiena), `${nomeOp} · rata mutuo${count < mesiPre ? " (pre-ammortamento)" : ""}`);
        cur = addMonths(cur, 1); count++;
      }
    } else if (!prestito.rimborsoAScadenza) {
      // Prestito ponte con interessi pagati mensilmente fino alla chiusura.
      const rata = importoPrestito * tassoMensile;
      let end = meseChiusura || start;
      if (ignoreSale) end = addMonths(start, horizonMonths);
      if (start && rata) {
        let cur = start, guard = 0;
        while (monthDiff(cur, end) >= 0 && guard < 400) {
          if (!giaRiconciliato(cur)) addFlow(flows, cur, -rata, `${nomeOp} · interessi prestito`);
          cur = addMonths(cur, 1); guard++;
        }
      }
    }
  }

  // Vendite: incasso lordo in entrata, rimborsi e imposte in uscita nello stesso mese.
  // Con il frazionamento ogni unità porta la sua quota proporzionale di rimborsi.
  vendite.forEach((v) => {
    const quota = ricavoTot > 0 ? v.importo / ricavoTot : 0;
    addFlow(flows, v.mese, v.importo, `${nomeOp} · incasso vendita${vendite.length > 1 ? ` — ${v.descrizione}` : ""}`);

    let uscite = 0;
    if (importoPrestito > 0) {
      uscite += importoPrestito * quota;
      if (prestito.rimborsoAScadenza) uscite += econ.interessiTotali * quota;
    }
    uscite += econ.imposteSocietarie * quota;
    uscite += econ.share.investitori * quota;      // restituzione del capitale dei soci
    uscite += econ.utileInvestitoriLordo * quota;  // quota di utile spettante ai soci
    if (uscite) addFlow(flows, v.mese, -uscite, `${nomeOp} · alla vendita: rimborsi, imposte e liquidazione soci`);
  });

  return flows;
}

// French-amortization monthly payment for a mortgage of `principal` at annual `rate` (%) over `months`.
function rataMutuo(principal, annualRatePct, months) {
  const P = Number(principal) || 0;
  const n = Number(months) || 0;
  const i = (Number(annualRatePct) || 0) / 100 / 12;
  if (P <= 0 || n <= 0) return 0;
  if (i === 0) return P / n;
  return (P * i) / (1 - Math.pow(1 + i, -n));
}

// Outstanding bank debt today. A bridge loan (flipping) keeps its full principal until the sale;
// an amortising mortgage is reduced by the capital already repaid, net of any pre-amortisation
// window during which only interest is paid.
function debitoResiduo(prestito, ammortizza, ym = currentYm()) {
  if (!prestito?.attivo) return 0;
  const P = Number(prestito.importo) || 0;
  if (P <= 0) return 0;
  const start = prestito.meseErogazione;
  if (start && monthDiff(start, ym) < 0) return 0; // non ancora erogato
  if (!ammortizza) return P;

  const durata = Math.max(0, Math.floor(Number(prestito.durataMesi) || 0));
  const mesiPre = Math.max(0, Math.min(Math.floor(Number(prestito.mesiPreammortamento) || 0), durata));
  const trascorsi = start ? monthDiff(start, ym) : 0;
  if (trascorsi <= mesiPre) return P;

  const n = durata - mesiPre;
  if (n <= 0) return 0;
  const k = Math.min(n, trascorsi - mesiPre);
  const i = (Number(prestito.tassoAnnuo) || 0) / 100 / 12;
  if (i === 0) return Math.max(0, P * (1 - k / n));
  const powN = Math.pow(1 + i, n), powK = Math.pow(1 + i, k);
  return Math.max(0, P * (powN - powK) / (powN - 1));
}

// Bank financing currently outstanding across active operations and directly-held rentals.
function finanziamentiAttivi(operations, rentals) {
  let totale = 0;
  (operations || [])
    .filter((o) => o.stato === "in_corso" || o.stato === "pipeline")
    .forEach((o) => { totale += debitoResiduo(o.copertura?.prestito, o.destinazione === "reddito"); });
  // I mutui degli immobili collegati a un'operazione sono già contati sopra.
  (rentals || []).forEach((r) => {
    if (r.operazioneId) return;
    totale += debitoResiduo(r.mutuo, true);
  });
  return totale;
}

// Third-party investor capital currently at work in operations that are still open.
function capitaleInvestitoriAttivo(operations) {
  return (operations || [])
    .filter((o) => o.stato === "in_corso" || o.stato === "pipeline")
    .reduce((s, o) => s + fundingShare(o).investitori, 0);
}


// When `operazione` is provided the property was bought through an operation flagged "a reddito":
// acquisition costs and the mortgage are taken from there, so the letting sheet measures the real
// monthly outflows against the rental income without the data being entered twice.
// Resolves the operation a rental was bought through, if any.
const opDiRental = (rental, operations) =>
  rental?.operazioneId ? ((operations || []).find((o) => o.id === rental.operazioneId) || null) : null;

function rentalEconomics(rental, operazione = null) {
  const r = rental || {};
  const ric = r.ricavi || {};

  let ricavoAnnuoLordo = 0;
  if (r.tipoAffitto === "lungo") {
    ricavoAnnuoLordo = (Number(ric.canoneMensile) || 0) * 12;
  } else {
    const tariffa = Number(ric.tariffaNotte) || 0;
    const occ = Math.min(100, Math.max(0, Number(ric.occupazionePct) || 0));
    ricavoAnnuoLordo = tariffa * 365 * (occ / 100);
  }

  const costiRicorrentiAnnui = (r.costiRicorrenti || []).reduce((s, c) => s + (Number(c.importoMensile) || 0) * 12, 0);

  // Il finanziamento può arrivare dall'operazione collegata (mutuo) o essere inserito qui.
  const prestitoOp = operazione?.copertura?.prestito?.attivo ? operazione.copertura.prestito : null;
  const mutuo = prestitoOp || (r.mutuo?.attivo ? r.mutuo : null);
  const collegato = !!operazione;

  const durataTot = Math.max(0, Math.floor(Number(mutuo?.durataMesi) || 0));
  const mesiPre = mutuo ? Math.max(0, Math.min(Math.floor(Number(mutuo.mesiPreammortamento) || 0), durataTot)) : 0;
  const mesiAmmortamento = Math.max(0, durataTot - mesiPre);
  const rataPre = mutuo ? (Number(mutuo.importo) || 0) * (Number(mutuo.tassoAnnuo) || 0) / 100 / 12 : 0;
  const rataPiena = mutuo ? rataMutuo(mutuo.importo, mutuo.tassoAnnuo, mesiAmmortamento) : 0;
  // Headline metrics use the fully amortising instalment: it is the sustained, steady-state cost.
  const rataMensile = rataPiena;
  const rataMutuoAnnua = rataMensile * 12;

  const investimentoProprio = (r.costiUnaTantum || []).reduce((s, c) => s + (Number(c.importo) || 0), 0);
  const investimentoOperazione = collegato
    ? (operazione.costiUnaTantum || []).reduce((s, c) => s + (Number(c.importo) || 0), 0)
    : 0;
  const investimentoTotale = investimentoProprio + investimentoOperazione;
  const capitaleFinanziato = mutuo ? (Number(mutuo.importo) || 0) : 0;
  const capitaleProprio = Math.max(0, investimentoTotale - capitaleFinanziato);

  const cashFlowAnnuoNetto = ricavoAnnuoLordo - costiRicorrentiAnnui - rataMutuoAnnua;
  // Net operating income (before financing) — the return the asset produces on its own.
  const noiAnnuo = ricavoAnnuoLordo - costiRicorrentiAnnui;

  const rendimentoLordo = investimentoTotale > 0 ? (ricavoAnnuoLordo / investimentoTotale) * 100 : 0;
  const rendimentoNetto = investimentoTotale > 0 ? (noiAnnuo / investimentoTotale) * 100 : 0;
  const cashOnCash = capitaleProprio > 0 ? (cashFlowAnnuoNetto / capitaleProprio) * 100 : 0;

  return {
    ricavoAnnuoLordo, costiRicorrentiAnnui, rataMensile, rataMutuoAnnua, noiAnnuo,
    investimentoTotale, capitaleFinanziato, capitaleProprio,
    cashFlowAnnuoNetto, rendimentoLordo, rendimentoNetto, cashOnCash,
    // Pre-amortisation detail: during that window the cash flow is temporarily better.
    mesiPre, mesiAmmortamento, rataPre, rataPiena,
    cashFlowAnnuoPre: noiAnnuo - rataPre * 12,
    collegato, investimentoOperazione, investimentoProprio,
  };
}

// Monthly liquidity flows produced by a rental over the projection horizon:
// upfront costs at their date, recurring costs, mortgage instalments, and rental income
// from the start-of-letting month onwards.
// When a rental is linked to an operation (`operazioneId`), acquisition costs and the mortgage
// live in that operation and are already projected there: here we only add the letting side
// (management costs and rental income), otherwise they would be counted twice.
function rentalFlows(rental, horizonEndYm, cutoff, operazione = null) {
  const flows = {};
  const r = rental || {};
  const nomeR = r.nome || "Immobile a reddito";
  const collegato = !!operazione;

  if (!collegato) {
    (r.costiUnaTantum || []).forEach((c) => {
      if (!c.data) return;
      if (cutoff && c.data <= cutoff) return;
      addFlow(flows, ymdToYm(c.data), -(Number(c.importo) || 0), `${nomeR} · ${c.categoria || "costo"}`);
    });
  }

  (r.costiRicorrenti || []).forEach((c) => {
    const start = c.meseInizio || r.meseInizioLocazione;
    if (!start) return;
    const end = c.meseFine || horizonEndYm;
    let cur = start, guard = 0;
    while (monthDiff(cur, end) >= 0 && guard < 600) {
      if (!(cutoff && dueDateForMonth(cur, c.giornoAddebito) <= cutoff)) {
        addFlow(flows, cur, -(Number(c.importoMensile) || 0), `${nomeR} · ${c.categoria || "gestione"}`);
      }
      cur = addMonths(cur, 1); guard++;
    }
  });

  const mutuo = !collegato && r.mutuo?.attivo ? r.mutuo : null;
  if (mutuo && Number(mutuo.importo) > 0 && Number(mutuo.durataMesi) > 0) {
    const econ = rentalEconomics(r);
    const start = mutuo.meseErogazione || r.meseInizioLocazione;
    const durata = Math.floor(Number(mutuo.durataMesi) || 0);
    let cur = start, count = 0;
    while (count < durata && monthDiff(cur, horizonEndYm) >= 0) {
      // Interest only during the pre-amortisation window, full instalment afterwards.
      const rata = count < econ.mesiPre ? econ.rataPre : econ.rataPiena;
      if (!(cutoff && dueDateForMonth(cur, "") <= cutoff)) addFlow(flows, cur, -rata, `${nomeR} · rata mutuo`);
      cur = addMonths(cur, 1); count++;
    }
  }

  if (r.attivo) {
    const econ = rentalEconomics(r, operazione);
    const ricavoMensile = econ.ricavoAnnuoLordo / 12;
    if (ricavoMensile) {
      let cur = r.meseInizioLocazione || currentYm(), guard = 0;
      while (monthDiff(cur, horizonEndYm) >= 0 && guard < 600) {
        if (!(cutoff && dueDateForMonth(cur, "") <= cutoff)) addFlow(flows, cur, ricavoMensile, `${nomeR} · incasso affitti`);
        cur = addMonths(cur, 1); guard++;
      }
    }
  }

  return flows;
}

// Equity portfolio: per-position book value, market value and P/L (all converted to EUR),
// plus portfolio totals. Prices are entered manually; USD positions are converted with usdToEur.
function portfolioSummary(portafoglio, usdToEur) {
  const rate = Number(usdToEur) > 0 ? Number(usdToEur) : 1;
  let hasUsd = false;
  const titoli = (portafoglio?.titoli || []).map((t) => {
    const q = Number(t.quantita) || 0;
    const pc = Number(t.prezzoCarico) || 0;
    const pa = Number(t.prezzoAttuale) || 0;
    const valuta = t.valuta === "USD" ? "USD" : "EUR";
    if (valuta === "USD") hasUsd = true;
    const fx = valuta === "USD" ? rate : 1;
    // Native-currency values (for display) and EUR-converted values (for totals).
    const valoreCaricoNat = q * pc;
    const valoreAttualeNat = q * pa;
    const valoreCarico = valoreCaricoNat * fx;
    const valoreAttuale = valoreAttualeNat * fx;
    const pl = valoreAttuale - valoreCarico;
    const plPct = valoreCarico > 0 ? (pl / valoreCarico) * 100 : 0;
    return { ...t, valuta, q, pc, pa, valoreCaricoNat, valoreAttualeNat, valoreCarico, valoreAttuale, pl, plPct };
  });
  const investito = titoli.reduce((s, t) => s + t.valoreCarico, 0);
  const valoreAttuale = titoli.reduce((s, t) => s + t.valoreAttuale, 0);
  const pl = valoreAttuale - investito;
  const plPct = investito > 0 ? (pl / investito) * 100 : 0;
  titoli.forEach((t) => { t.peso = valoreAttuale > 0 ? (t.valoreAttuale / valoreAttuale) * 100 : 0; });
  return { titoli, investito, valoreAttuale, pl, plPct, hasUsd, rate };
}

// Fetches the daily USD→EUR rate from public, key-less, CORS-enabled endpoints (ECB-based).
// Returns { usdToEur, date } or throws. Tries Frankfurter first, then open.er-api as fallback.
// Extracts the first JSON array/object found in a model reply, tolerating code fences.
function extractJson(text, kind = "array") {
  if (!text) return null;
  const clean = String(text).replace(/```json/gi, "").replace(/```/g, "").trim();
  const open = kind === "array" ? "[" : "{";
  const close = kind === "array" ? "]" : "}";
  const start = clean.indexOf(open);
  const end = clean.lastIndexOf(close);
  if (start < 0 || end <= start) return null;
  try { return JSON.parse(clean.slice(start, end + 1)); } catch (e) { return null; }
}

// Calls the Anthropic API with the web_search tool, through the `claude-proxy` Edge Function
// (the API key stays server-side). This is the reliable route for data that market APIs won't
// serve to a browser (no CORS, or an API key that would be exposed in client code).
async function askClaudeJson(prompt, kind = "array") {
  const data = await callClaude({
    max_tokens: 1000,
    messages: [{ role: "user", content: prompt }],
    tools: [{ type: "web_search_20250305", name: "web_search" }],
  });
  const testo = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
  return extractJson(testo, kind);
}

// Looks up the latest available price for each holding that has a ticker.
async function fetchQuotesViaClaude(titoli) {
  const conTicker = (titoli || []).filter((t) => (t.ticker || "").trim());
  if (!conTicker.length) {
    return { error: "Nessun titolo ha un ticker. Inserisci il simbolo di borsa (es. AAPL, IWDA.AS) per poter cercare la quotazione." };
  }
  const lista = conTicker.map((t) => `${t.ticker.trim().toUpperCase()}${t.nome ? ` (${t.nome})` : ""}`).join(", ");
  const prompt = `Cerca sul web la quotazione più recente disponibile per questi strumenti finanziari: ${lista}.

Per ciascuno riporta l'ultimo prezzo noto, la valuta in cui è quotato e la data a cui il prezzo si riferisce.

Rispondi ESCLUSIVAMENTE con un array JSON, senza alcun testo prima o dopo e senza backtick, in questo formato esatto:
[{"ticker":"AAPL","prezzo":210.25,"valuta":"USD","data":"2026-07-24","fonte":"Nasdaq"}]

Regole: "prezzo" deve essere un numero senza separatori delle migliaia e con il punto come separatore decimale. "valuta" deve essere il codice a 3 lettere (EUR, USD, ...). Se per uno strumento non trovi un prezzo affidabile, metti "prezzo": null. Non inventare prezzi.`;

  const quotes = await askClaudeJson(prompt, "array");
  if (!Array.isArray(quotes)) return { error: "Non sono riuscito a interpretare la risposta. Riprova." };
  return { quotes };
}

// Fallback for the FX rate when the direct call to a public endpoint is blocked.
async function fetchFxViaClaude() {
  const obj = await askClaudeJson(
    `Cerca sul web il tasso di cambio più recente da dollaro USA a euro (quanti euro vale 1 USD).
Rispondi ESCLUSIVAMENTE con un oggetto JSON, senza testo prima o dopo e senza backtick, in questo formato:
{"usdToEur":0.9215,"data":"2026-07-24","fonte":"BCE"}
Il valore deve essere un numero con il punto come separatore decimale.`,
    "object"
  );
  const rate = Number(obj?.usdToEur);
  if (!Number.isFinite(rate) || rate <= 0) throw new Error("fx via claude non disponibile");
  return { usdToEur: rate, date: obj?.data || todayStr() };
}

async function fetchFxRate() {
  try {
    const res = await fetch("https://api.frankfurter.app/latest?from=USD&to=EUR");
    if (res.ok) {
      const data = await res.json();
      const rate = data?.rates?.EUR;
      if (rate) return { usdToEur: rate, date: data.date || todayStr() };
    }
  } catch (e) { /* fall through to backup */ }

  try {
    const res2 = await fetch("https://open.er-api.com/v6/latest/USD");
    if (res2.ok) {
      const data2 = await res2.json();
      const rate2 = data2?.rates?.EUR;
      if (rate2) {
        const date2 = data2?.time_last_update_utc ? new Date(data2.time_last_update_utc).toISOString().slice(0, 10) : todayStr();
        return { usdToEur: rate2, date: date2 };
      }
    }
  } catch (e) { /* fall through to Claude */ }

  // Last resort: ask Claude to look it up (works even when external domains are blocked).
  return fetchFxViaClaude();
}

// Unsettled receivables (money coming in) and payables (money going out), each on its single due date.
function partiteFlows(partite) {
  const flows = {};
  (partite || []).forEach((p) => {
    if (p.saldato) return;
    if (!p.data) return;
    const amt = Number(p.importo) || 0;
    if (!amt) return;
    addFlow(flows, ymdToYm(p.data), p.tipo === "debito" ? -amt : amt, `${p.tipo === "debito" ? "Debito verso" : "Credito da"} ${p.controparte || "controparte"}${p.descrizione ? ` — ${p.descrizione}` : ""}`);
  });
  return flows;
}

// Company-wide fixed costs (salaries, office, etc.) — independent of any single operation.
// Recurring items with no end month are assumed to continue for the whole projection horizon.
function costiFissiFlows(costiFissi, horizonEndYm, cutoff) {
  const flows = {};
  (costiFissi?.unaTantum || []).forEach((c) => {
    if (!c.data) return;
    if (cutoff && c.data <= cutoff) return;
    addFlow(flows, ymdToYm(c.data), -(Number(c.importo) || 0), `Costo fisso · ${c.descrizione || c.categoria || "una tantum"}`);
  });
  (costiFissi?.ricorrenti || []).forEach((c) => {
    const start = c.meseInizio;
    if (!start) return;
    const end = c.meseFine || horizonEndYm;
    let cur = start, guard = 0;
    while (monthDiff(cur, end) >= 0 && guard < 400) {
      if (!(cutoff && dueDateForMonth(cur, c.giornoAddebito) <= cutoff)) {
        addFlow(flows, cur, -(Number(c.importoMensile) || 0), `Costo fisso · ${c.descrizione || c.categoria || "ricorrente"}`);
      }
      cur = addMonths(cur, 1);
      guard++;
    }
  });
  return flows;
}

// Aggregates every source of cash movement over a given list of months, keeping gross inflows
// and outflows apart. `cutoff` suppresses what is already inside the reconciled balances;
// pass null to get the plain plan (used by the calendar-year chart, which is a planning view).
function aggregateFlows(liquidity, operations, rentals, months, cutoff, opts = {}) {
  const acc = {};
  months.forEach((m) => (acc[m] = { in: 0, out: 0, voci: [] }));
  const merge = (flows) => {
    Object.entries(flows || {}).forEach(([ym, cell]) => {
      if (!acc[ym]) return;
      acc[ym].in += cell.in || 0;
      acc[ym].out += cell.out || 0;
      if (cell.voci) acc[ym].voci.push(...cell.voci);
    });
  };

  const movFlows = {};
  (liquidity.movimenti || []).forEach((mv) => {
    if (cutoff && mv.data && mv.data <= cutoff) return;
    addFlow(movFlows, ymdToYm(mv.data), Number(mv.importo) || 0, `Movimento extra · ${mv.descrizione || "senza descrizione"}`);
  });
  merge(movFlows);

  merge(partiteFlows(liquidity.partite));

  const horizonEnd = months[months.length - 1];
  merge(costiFissiFlows(liquidity.costiFissi, horizonEnd, cutoff));

  (operations || [])
    .filter((o) => o.stato === "in_corso" || o.stato === "pipeline")
    .forEach((op) => {
      const ignoreSale = opts.ignoreSaleForOpId === op.id;
      merge(operationQGFlows(op, liquidity, { ignoreSale, horizonMonths: months.length, cutoff }));
    });

  // Immobili a reddito: gestione e affitti (acquisto e mutuo restano sull'operazione collegata).
  (rentals || []).forEach((rental) => {
    merge(rentalFlows(rental, horizonEnd, cutoff, opDiRental(rental, operations)));
  });

  return acc;
}

function computeGlobalProjection(liquidity, operations, rentals, monthsAhead, opts = {}) {
  const start = currentYm();
  const months = Array.from({ length: monthsAhead }, (_, i) => addMonths(start, i));
  // Reconciliation cutoff: anything dated on or before it is already inside the account
  // balances, so re-projecting it would double-count. Debts/credits are deliberately exempt:
  // they carry their own "saldato" flag to say whether they have actually been settled.
  const acc = aggregateFlows(liquidity, operations, rentals, months, liquidity.dataRiferimento || null, opts);

  let running = totaleDisponibile(liquidity);
  return months.map((m) => {
    const cell = acc[m];
    const net = cell.in - cell.out;
    running += net;
    // voci ordinate per impatto, così il dettaglio del mese si legge dall'importo più rilevante
    const voci = (cell.voci || []).slice().sort((a, b) => Math.abs(b.importo) - Math.abs(a.importo));
    return { month: m, label: monthLabel(m), entrate: cell.in, uscite: cell.out, net, balance: running, voci };
  });
}

// Planned inflows/outflows for each month of a calendar year. No cutoff here: the point is to
// see the shape of the year as planned, including months already settled.
function computeAnnualFlows(liquidity, operations, rentals, anno) {
  const months = Array.from({ length: 12 }, (_, i) => `${anno}-${String(i + 1).padStart(2, "0")}`);
  const acc = aggregateFlows(liquidity, operations, rentals, months, null);
  return months.map((m, i) => ({
    month: m,
    mese: MESI_BREVI[i],
    entrate: acc[m].in,
    uscite: acc[m].out,
    netto: acc[m].in - acc[m].out,
    voci: (acc[m].voci || []).slice().sort((a, b) => Math.abs(b.importo) - Math.abs(a.importo)),
  }));
}

function analyzeSustainability(op, liquidity, operations, rentals = []) {
  const soglia = Number(liquidity.soglioMinimo) || 0;
  const econ = operationEconomics(op, liquidity);
  // Un'operazione a reddito non si valuta sul margine di vendita: la sua redditività
  // sta nella scheda dell'immobile, qui conta solo che la cassa regga.
  const profittevole = econ.aReddito ? true : econ.marginNetto > 0;

  const monthsAhead = 72;
  const rows = computeGlobalProjection(liquidity, operations, rentals, monthsAhead);
  const activeStart = op.meseInizio;
  const activeEnd = econ.meseChiusura || addMonths(op.meseInizio, 12);
  const activeRows = rows.filter((r) => monthDiff(activeStart, r.month) >= 0 && monthDiff(r.month, activeEnd) >= 0);
  const minRow = activeRows.reduce((min, r) => (r.balance < min.balance ? r : min), activeRows[0] || { balance: Infinity, month: activeStart });

  const sostenibileConLiquidita = econ.share.gap <= 0 && minRow.balance >= soglia;
  const deficitLiquidita = Math.max(0, soglia - minRow.balance);

  // months sustainable without ever selling
  const rowsNoSale = computeGlobalProjection(liquidity, operations, rentals, monthsAhead, { ignoreSaleForOpId: op.id });
  const fromStartIdx = rowsNoSale.findIndex((r) => monthDiff(activeStart, r.month) >= 0);
  let mesiSostenibili = Infinity;
  for (let i = Math.max(0, fromStartIdx); i < rowsNoSale.length; i++) {
    if (rowsNoSale[i].balance < soglia) {
      mesiSostenibili = monthDiff(activeStart, rowsNoSale[i].month);
      break;
    }
  }

  let verdict = "ok";
  if (!profittevole) verdict = "sconsigliata";
  else if (econ.share.gap > 0) verdict = "copertura_insufficiente";
  else if (!sostenibileConLiquidita) verdict = "serve_finanziamento";

  return {
    ...econ, profittevole,
    minRow, sostenibileConLiquidita, deficitLiquidita, mesiSostenibili, verdict, soglia,
  };
}

// Aggregates every investor (matched by name, case/space-insensitive) across all operations,
// splitting amounts into: still active (pipeline/in corso — projected), completed (already
// liquidated) and abandoned (capital at risk / needs manual reconciliation).
function getInvestorsSummary(operations, liquidity) {
  const map = new Map();
  operations.forEach((op) => {
    if (!op.copertura?.investitori?.attivo) return;
    const econ = operationEconomics(op, liquidity);
    econ.investitoriDettaglio.forEach((inv) => {
      const nomeRaw = (inv.nome || "").trim();
      const nome = nomeRaw || "Investitore senza nome";
      const key = nome.toLowerCase();
      if (!map.has(key)) map.set(key, { nome, operazioni: [] });
      map.get(key).operazioni.push({
        opId: op.id,
        opNome: op.nome || "Senza nome",
        stato: op.stato,
        meseVenditaPrevista: meseChiusuraOperazione(op),
        importo: inv.importo,
        quotaPct: inv.quotaPct,
        utileLordo: inv.utileLordo,
        impostaSostitutiva: inv.impostaSostitutiva,
        utileNetto: inv.utileNetto,
        totaleLiquidato: inv.totaleLiquidato,
      });
    });
  });

  const investors = Array.from(map.values()).map((entry) => {
    const attive = entry.operazioni.filter((o) => o.stato === "in_corso" || o.stato === "pipeline");
    const completate = entry.operazioni.filter((o) => o.stato === "completata");
    const abbandonate = entry.operazioni.filter((o) => o.stato === "abbandonata");
    return {
      nome: entry.nome,
      operazioni: entry.operazioni,
      numeroOperazioni: entry.operazioni.length,
      capitaleAttivo: attive.reduce((s, o) => s + o.importo, 0),
      daRestituireProiettato: attive.reduce((s, o) => s + o.totaleLiquidato, 0),
      utileNettoProiettato: attive.reduce((s, o) => s + o.utileNetto, 0),
      // Utile prima della sostitutiva del 26%: è la base se l'investitore lascia correre l'utile
      // invece di prelevarlo (l'imposta si applica al momento della liquidazione).
      utileLordoProiettato: attive.reduce((s, o) => s + o.utileLordo, 0),
      impostaProiettata: attive.reduce((s, o) => s + o.impostaSostitutiva, 0),
      capitaleCompletate: completate.reduce((s, o) => s + o.importo, 0),
      incassatoCompletate: completate.reduce((s, o) => s + o.totaleLiquidato, 0),
      capitaleAbbandonate: abbandonate.reduce((s, o) => s + o.importo, 0),
      capitaleTotaleStorico: entry.operazioni.reduce((s, o) => s + o.importo, 0),
    };
  });
  investors.sort((a, b) => (b.capitaleAttivo + b.capitaleCompletate) - (a.capitaleAttivo + a.capitaleCompletate));

  const totals = investors.reduce((acc, inv) => ({
    capitaleAttivo: acc.capitaleAttivo + inv.capitaleAttivo,
    daRestituireProiettato: acc.daRestituireProiettato + inv.daRestituireProiettato,
    incassatoCompletate: acc.incassatoCompletate + inv.incassatoCompletate,
    capitaleAbbandonate: acc.capitaleAbbandonate + inv.capitaleAbbandonate,
  }), { capitaleAttivo: 0, daRestituireProiettato: 0, incassatoCompletate: 0, capitaleAbbandonate: 0 });

  return { investors, totals };
}

// Collects every one-off expense (operations + company fixed costs) into a single day-level
// schedule, sorted by date. Recurring costs are excluded on purpose — they have no single due day.
// Collects every one-off expense (operations + company fixed costs) into a single day-level
// schedule, sorted by date. Recurring costs are excluded on purpose — they have no single due day.
// Only operations that still affect cash (pipeline / in corso) are listed, matching the projection:
// an abandoned or closed operation will not produce those payments.
function buildSpeseSchedule(operations, liquidity) {
  const items = [];
  (operations || [])
    .filter((op) => op.stato === "in_corso" || op.stato === "pipeline")
    .forEach((op) => {
    (op.costiUnaTantum || []).forEach((c) => {
      if (!c.data) return;
      items.push({
        id: c.id,
        data: c.data,
        importo: Number(c.importo) || 0,
        categoria: c.categoria || "Costo",
        origine: op.nome || "Operazione senza nome",
        tipo: "operazione",
        opId: op.id,
        stato: op.stato,
      });
    });
  });
  (liquidity?.costiFissi?.unaTantum || []).forEach((c) => {
    if (!c.data) return;
    items.push({
      id: c.id,
      data: c.data,
      importo: Number(c.importo) || 0,
      categoria: c.categoria || "Costo",
      origine: c.descrizione || "Costo fisso aziendale",
      tipo: "fisso",
    });
  });
  items.sort((a, b) => (a.data < b.data ? -1 : a.data > b.data ? 1 : 0));
  return items;
}

/* ============================================================
   CSV BANK IMPORT (Unicredit statement)
   ============================================================ */
// Splits a CSV line respecting double-quoted fields (which may contain the delimiter).
function splitCsvLine(line, delim) {
  const out = [];
  let cur = "", inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
      else inQ = !inQ;
    } else if (ch === delim && !inQ) {
      out.push(cur); cur = "";
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out.map((s) => s.trim());
}

// Parses a European-formatted amount: "1.234,56" or "-1.234,56" or "1234,56" → number.
function parseEuroAmount(raw) {
  if (raw == null) return NaN;
  // Excel cells arrive already as numbers — use them directly.
  if (typeof raw === "number") return Number.isFinite(raw) ? raw : NaN;
  let s = String(raw).trim().replace(/[€\s]/g, "");
  if (!s) return NaN;
  // Detect sign in parentheses e.g. (1.234,56)
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  // If both separators present, the last one is the decimal separator.
  const lastComma = s.lastIndexOf(",");
  const lastDot = s.lastIndexOf(".");
  if (lastComma > -1 && lastDot > -1) {
    if (lastComma > lastDot) { s = s.replace(/\./g, "").replace(",", "."); }
    else { s = s.replace(/,/g, ""); }
  } else if (lastComma > -1) {
    s = s.replace(/\./g, "").replace(",", ".");
  }
  const n = Number(s);
  if (Number.isNaN(n)) return NaN;
  return neg ? -n : n;
}

// Parses a date cell into "YYYY-MM-DD". Handles JS Date objects (from Excel), plus
// dd/mm/yyyy, dd-mm-yyyy, dd.mm.yyyy and yyyy-mm-dd string formats.
function parseBankDate(raw) {
  if (raw == null || raw === "") return null;
  // Excel cells read with cellDates arrive as JS Date objects.
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
    return `${raw.getFullYear()}-${String(raw.getMonth() + 1).padStart(2, "0")}-${String(raw.getDate()).padStart(2, "0")}`;
  }
  const s = String(raw).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})/);
  if (m) {
    let [, d, mo, y] = m;
    if (y.length === 2) y = "20" + y;
    return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  return null;
}

const norm = (s) => (s || "").toLowerCase().replace(/[^a-z]/g, "");

// Core parser working on a matrix of rows (each row = array of raw cell values).
// Shared by both the CSV path (cells are strings) and the Excel path (cells may be
// numbers or Date objects). Returns { movimenti, saldo, saldoData, meta } or { error }.
function parseBankRows(rows) {
  const grid = (rows || []).filter((r) => Array.isArray(r) && r.some((c) => c != null && String(c).trim() !== ""));
  if (grid.length < 2) return { error: "Il file non contiene righe sufficienti." };

  // Find the header row: the first row containing a recognizable date-ish column name.
  let headerIdx = 0;
  for (let i = 0; i < Math.min(grid.length, 15); i++) {
    const cells = grid[i].map((c) => norm(String(c ?? "")));
    if (cells.some((c) => c.includes("data") || c.includes("date"))) { headerIdx = i; break; }
  }
  const header = grid[headerIdx].map((c) => String(c ?? ""));
  const H = header.map(norm);

  const findCol = (preds) => H.findIndex((h) => preds.some((p) => h.includes(p)));
  let idxData = findCol(["datavaluta", "datacontabile", "dataoperazione", "data", "date"]);
  const idxDesc = findCol(["descrizione", "causale", "dettagli", "description"]);
  const idxImporto = findCol(["importo", "amount"]);
  const idxDare = findCol(["dare", "uscite", "addebiti", "debit"]);
  const idxAvere = findCol(["avere", "entrate", "accrediti", "credit"]);
  const idxSaldo = findCol(["saldo", "balance"]);

  if (idxData < 0) idxData = 0;
  const hasSingleAmount = idxImporto >= 0;
  const hasSplitAmount = idxDare >= 0 || idxAvere >= 0;
  if (!hasSingleAmount && !hasSplitAmount) {
    return { error: "Non ho trovato una colonna importo (o Dare/Avere) nell'intestazione del file." };
  }

  const movimenti = [];
  let lastSaldo = null, lastSaldoData = null;
  for (let i = headerIdx + 1; i < grid.length; i++) {
    const cells = grid[i];
    if (cells.length <= idxData) continue;
    const data = parseBankDate(cells[idxData]);
    if (!data) continue;

    let importo = NaN;
    const rawImporto = hasSingleAmount ? cells[idxImporto] : null;
    if (hasSingleAmount && rawImporto != null && String(rawImporto).trim() !== "") {
      importo = parseEuroAmount(rawImporto);
    }
    if (Number.isNaN(importo) && hasSplitAmount) {
      const dare = idxDare >= 0 ? parseEuroAmount(cells[idxDare]) : NaN;
      const avere = idxAvere >= 0 ? parseEuroAmount(cells[idxAvere]) : NaN;
      const d = Number.isNaN(dare) ? 0 : Math.abs(dare);
      const a = Number.isNaN(avere) ? 0 : Math.abs(avere);
      if (d === 0 && a === 0) importo = NaN;
      else importo = a - d;
    }
    if (Number.isNaN(importo)) continue;

    const descrizione = idxDesc >= 0 ? String(cells[idxDesc] ?? "") : "";
    movimenti.push({ id: uid(), data, importo: String(importo), descrizione: descrizione.trim().slice(0, 140) });

    if (idxSaldo >= 0) {
      const sal = parseEuroAmount(cells[idxSaldo]);
      if (!Number.isNaN(sal)) { lastSaldo = sal; lastSaldoData = data; }
    }
  }

  if (movimenti.length === 0) return { error: "Nessun movimento valido trovato nel file. Controlla che sia l'estratto conto della banca." };

  return { movimenti, saldo: lastSaldo, saldoData: lastSaldoData, meta: { colonne: header, righe: movimenti.length } };
}

// Reads a CSV export: auto-detects the delimiter, splits into a row matrix, then delegates
// to the shared row parser.
function parseBankCSV(text) {
  const clean = text.replace(/^\uFEFF/, "");
  const rawLines = clean.split(/\r\n|\n|\r/).filter((l) => l.trim().length > 0);
  if (rawLines.length < 2) return { error: "Il file non contiene righe sufficienti." };
  const first = rawLines[0];
  const candidates = [";", ",", "\t"];
  const delim = candidates.map((d) => ({ d, n: first.split(d).length })).sort((a, b) => b.n - a.n)[0].d;
  const rows = rawLines.map((l) => splitCsvLine(l, delim));
  const res = parseBankRows(rows);
  if (res.meta) res.meta.delimiter = delim;
  return res;
}

// Reads an .xls / .xlsx export via SheetJS: converts the first sheet into a row matrix
// (dates as JS Date objects, amounts as numbers) and delegates to the shared row parser.
function parseBankXLSX(arrayBuffer) {
  try {
    const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
    const firstSheet = wb.SheetNames[0];
    if (!firstSheet) return { error: "Il file Excel non contiene fogli leggibili." };
    const ws = wb.Sheets[firstSheet];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, blankrows: false, defval: "" });
    const res = parseBankRows(rows);
    if (res.meta) res.meta.foglio = firstSheet;
    return res;
  } catch (e) {
    return { error: "Impossibile leggere il file Excel: potrebbe essere danneggiato o protetto." };
  }
}

// Dispatches to the right reader based on the file extension, reading the file with the
// appropriate method (text for CSV, binary ArrayBuffer for Excel).
async function parseBankFile(file) {
  const name = (file.name || "").toLowerCase();
  if (name.endsWith(".xls") || name.endsWith(".xlsx") || name.endsWith(".xlsm")) {
    const buf = await file.arrayBuffer();
    return parseBankXLSX(buf);
  }
  if (name.endsWith(".csv") || name.endsWith(".txt")) {
    const text = await file.text();
    return parseBankCSV(text);
  }
  // Unknown extension: sniff by trying Excel first (binary signature), then CSV.
  const buf = await file.arrayBuffer();
  const head = new Uint8Array(buf.slice(0, 8));
  // XLSX = ZIP ("PK"), legacy XLS = OLE2 (0xD0 0xCF)
  const isZip = head[0] === 0x50 && head[1] === 0x4b;
  const isOle = head[0] === 0xd0 && head[1] === 0xcf;
  if (isZip || isOle) return parseBankXLSX(buf);
  const text = new TextDecoder("utf-8").decode(buf);
  return parseBankCSV(text);
}

/* ============================================================
   SMALL UI PIECES
   ============================================================ */
function StatCard({ icon: Icon, label, value, sub, tone = "neutral", extra = null }) {
  const toneColor = tone === "positive" ? "var(--positive)" : tone === "negative" ? "var(--negative)" : tone === "warning" ? "var(--warning)" : "var(--accent)";
  return (
    <div className="qg-panel" style={{ padding: "18px 20px", flex: 1, minWidth: 200 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
        <Icon size={15} color={toneColor} />
        <span className="qg-label" style={{ marginBottom: 0 }}>{label}</span>
        {extra && <span style={{ marginLeft: "auto" }}>{extra}</span>}
      </div>
      <div className="qg-mono qg-display" style={{ fontSize: 26, fontWeight: 600, color: "var(--text)", lineHeight: 1.1 }}>{value}</div>
      {sub && <div style={{ fontSize: 12, color: "var(--text-muted)", marginTop: 6 }}>{sub}</div>}
    </div>
  );
}

function VerdictBadge({ verdict }) {
  const map = {
    ok: { label: "Sostenibile", color: "var(--positive)", bg: "var(--positive-soft)", Icon: CheckCircle2 },
    serve_finanziamento: { label: "Serve più capitale", color: "var(--warning)", bg: "var(--warning-soft)", Icon: AlertTriangle },
    copertura_insufficiente: { label: "Copertura incompleta", color: "var(--warning)", bg: "var(--warning-soft)", Icon: AlertTriangle },
    sconsigliata: { label: "Sconsigliata", color: "var(--negative)", bg: "var(--negative-soft)", Icon: ShieldAlert },
  };
  const v = map[verdict] || map.ok;
  return (
    <span className="qg-badge" style={{ color: v.color, background: v.bg }}>
      <v.Icon size={12} /> {v.label}
    </span>
  );
}

function StatoBadge({ stato }) {
  const s = statoInfo(stato);
  return (
    <span className="qg-badge" style={{ color: s.color, background: "rgba(255,255,255,0.05)", border: `1px solid ${s.color}55` }}>
      {s.label}
    </span>
  );
}

function Field({ label, children }) {
  return (
    <div>
      <label className="qg-label">{label}</label>
      {children}
    </div>
  );
}

function NavItem({ icon: Icon, label, active, onClick, count }) {
  return (
    <div
      onClick={onClick}
      style={{
        display: "flex", alignItems: "center", gap: 10, padding: "10px 14px", borderRadius: 8,
        cursor: "pointer", fontSize: 13.5, fontWeight: 500,
        color: active ? "var(--accent-strong)" : "var(--text-muted)",
        background: active ? "var(--accent-soft)" : "transparent",
        transition: "all .15s",
      }}
    >
      <Icon size={16} />
      <span style={{ flex: 1 }}>{label}</span>
      {count != null && (
        <span className="qg-mono" style={{ fontSize: 11, color: active ? "var(--accent)" : "var(--text-dim)" }}>{count}</span>
      )}
    </div>
  );
}

/* Custom tooltip for the liquidity chart */
function LiquidityTooltip({ active, payload, label, valoreTitoli = 0 }) {
  if (!active || !payload || !payload.length) return null;
  const row = payload[0].payload;
  return (
    <div className="qg-panel qg-mono" style={{ padding: "10px 12px", fontSize: 12, border: "1px solid var(--border)" }}>
      <div style={{ color: "var(--text-muted)", marginBottom: 4 }}>{monthLabelLong(row.month)}</div>
      <div style={{ color: "var(--text)", fontWeight: 600 }}>Saldo: {fmtEUR(row.balance)}</div>
      <div style={{ color: row.net >= 0 ? "var(--positive)" : "var(--negative)" }}>Movimento: {fmtEURSigned(row.net)}</div>
      {valoreTitoli > 0 && <div style={{ color: "var(--text-dim)", marginTop: 4, paddingTop: 4, borderTop: "1px solid var(--border-soft)" }}>Con titoli: {fmtEUR(row.balance + valoreTitoli)}</div>}
    </div>
  );
}

/* ============================================================
   OPERATION EDITOR
   ============================================================ */
function OperationEditor({ operation, liquidity, allInvestorNames, onSave, onCancel, onDelete }) {
  const [op, setOp] = useState(operation);

  const update = (patch) => setOp((prev) => ({ ...prev, ...patch }));
  const updateCopertura = (patch) => setOp((prev) => ({ ...prev, copertura: { ...prev.copertura, ...patch } }));
  const updatePrestito = (patch) => setOp((prev) => ({ ...prev, copertura: { ...prev.copertura, prestito: { ...prev.copertura.prestito, ...patch } } }));
  const updateInvestitori = (patch) => setOp((prev) => ({ ...prev, copertura: { ...prev.copertura, investitori: { ...prev.copertura.investitori, ...patch } } }));

  // Un'entrata di capitale datata nel mese della riconciliazione viene comunque proiettata:
  // se quel denaro è già sul conto verrebbe contato due volte, quindi lo segnaliamo.
  const meseRiconciliazione = liquidity?.dataRiferimento ? ymdToYm(liquidity.dataRiferimento) : null;
  const capitaleDaVerificare = (ym) => !!(meseRiconciliazione && ym && ym === meseRiconciliazione && dueDateForMonth(ym, "") > liquidity.dataRiferimento);
  const AvvisoCapitale = ({ ym }) => capitaleDaVerificare(ym) ? (
    <div style={{ fontSize: 10.5, color: "var(--warning)", marginTop: 4, lineHeight: 1.35 }}>
      Cade nel mese dei saldi aggiornati: viene proiettato come entrata. Se questo denaro è già sul conto, indica un mese precedente.
    </div>
  ) : null;

  const addVendita = () => update({ vendite: [...(op.vendite || []), { id: uid(), descrizione: "", importo: "", mese: op.meseVenditaPrevista || currentYm() }] });
  const updVendita = (id, patch) => update({ vendite: (op.vendite || []).map((v) => (v.id === id ? { ...v, ...patch } : v)) });
  const delVendita = (id) => update({ vendite: (op.vendite || []).filter((v) => v.id !== id) });

  const addCostoUnaTantum = () => update({ costiUnaTantum: [...op.costiUnaTantum, { id: uid(), categoria: CAT_ONEOFF[0], importo: "", data: toDateStr(op.meseInizio) || todayStr() }] });
  const updCostoUnaTantum = (id, patch) => update({ costiUnaTantum: op.costiUnaTantum.map((c) => (c.id === id ? { ...c, ...patch } : c)) });
  const delCostoUnaTantum = (id) => update({ costiUnaTantum: op.costiUnaTantum.filter((c) => c.id !== id) });

  const addCostoRicorrente = () => update({ costiRicorrenti: [...op.costiRicorrenti, { id: uid(), categoria: CAT_RICORRENTE[0], importoMensile: "", meseInizio: op.meseInizio, meseFine: "" }] });
  const updCostoRicorrente = (id, patch) => update({ costiRicorrenti: op.costiRicorrenti.map((c) => (c.id === id ? { ...c, ...patch } : c)) });
  const delCostoRicorrente = (id) => update({ costiRicorrenti: op.costiRicorrenti.filter((c) => c.id !== id) });

  const addInvestitore = () => updateInvestitori({ lista: [...(op.copertura.investitori.lista || []), { id: uid(), nome: "", importo: "", meseVersamento: op.meseInizio || currentYm() }] });
  const updInvestitore = (id, patch) => updateInvestitori({ lista: op.copertura.investitori.lista.map((i) => (i.id === id ? { ...i, ...patch } : i)) });
  const delInvestitore = (id) => updateInvestitori({ lista: op.copertura.investitori.lista.filter((i) => i.id !== id) });

  const share = useMemo(() => fundingShare(op), [op]);
  const econ = useMemo(() => operationEconomics(op, liquidity), [op, liquidity]);
  const canSave = op.nome.trim().length > 0 && op.meseInizio;

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "16px 24px", borderBottom: "1px solid var(--border-soft)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button className="qg-btn qg-btn-ghost" onClick={onCancel}><ChevronLeft size={16} /> Indietro</button>
          <span className="qg-display" style={{ fontSize: 18, fontWeight: 600 }}>{operation.nome ? "Modifica operazione" : "Nuova operazione"}</span>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          {onDelete && (
            <button className="qg-btn qg-btn-ghost qg-btn-danger" onClick={() => onDelete(op.id)}><Trash2 size={14} /> Elimina</button>
          )}
          <button className="qg-btn qg-btn-primary" disabled={!canSave} onClick={() => onSave(op)}><Save size={14} /> Salva operazione</button>
        </div>
      </div>

      <div className="qg-scroll" style={{ flex: 1, overflowY: "auto", padding: "24px", display: "flex", flexDirection: "column", gap: 22 }}>
        {/* Info base */}
        <div className="qg-panel" style={{ padding: 20 }}>
          <div className="qg-label" style={{ fontSize: 12, marginBottom: 14, color: "var(--accent)" }}>Anagrafica</div>
          <div style={{ display: "grid", gridTemplateColumns: "2fr 2fr 1fr", gap: 14, marginBottom: 14 }}>
            <Field label="Nome operazione"><input className="qg-input" style={{ fontFamily: "Inter" }} value={op.nome} onChange={(e) => update({ nome: e.target.value })} placeholder="Es. Via Torino 12" /></Field>
            <Field label="Indirizzo / descrizione"><input className="qg-input" style={{ fontFamily: "Inter" }} value={op.indirizzo} onChange={(e) => update({ indirizzo: e.target.value })} placeholder="Milano, zona Navigli — trilocale" /></Field>
            <Field label="Stato">
              <select className="qg-input" value={op.stato} onChange={(e) => update({ stato: e.target.value })}>
                {STATI.map((s) => <option key={s.key} value={s.key}>{s.label}</option>)}
              </select>
            </Field>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14, marginBottom: 14 }}>
            <Field label="Destinazione">
              <select className="qg-input" value={op.destinazione} onChange={(e) => update({ destinazione: e.target.value })}>
                <option value="flipping">Flipping (rivendita)</option>
                <option value="reddito">A reddito (tenuto e affittato)</option>
              </select>
            </Field>
            <Field label="Mese di inizio"><input type="month" className="qg-input" value={op.meseInizio} onChange={(e) => update({ meseInizio: e.target.value })} /></Field>
            {op.destinazione === "flipping" ? (
              <Field label="Tipo di vendita">
                <select className="qg-input" value={op.frazionamento ? "fraz" : "unica"} onChange={(e) => update({ frazionamento: e.target.value === "fraz" })}>
                  <option value="unica">Vendita unica</option>
                  <option value="fraz">Frazionamento (più unità)</option>
                </select>
              </Field>
            ) : <div />}
          </div>

          {op.destinazione === "reddito" ? (
            <div style={{ fontSize: 12.5, color: "var(--text-dim)", padding: "10px 12px", background: "var(--panel-2)", borderRadius: 7, border: "1px solid var(--border-soft)" }}>
              Operazione a reddito: qui inserisci i costi d'acquisto e il mutuo, così la proiezione di cassa è completa. Ricavi da affitto, costi di gestione e redditività si gestiscono poi in <b>Immobili a reddito</b>, collegando la scheda a questa operazione.
            </div>
          ) : op.frazionamento ? (
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <div className="qg-label" style={{ marginBottom: 0 }}>Unità in vendita</div>
                <button className="qg-btn qg-btn-ghost" onClick={addVendita}><Plus size={14} /> Aggiungi unità</button>
              </div>
              {(op.vendite || []).length === 0 && <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginBottom: 8 }}>Nessuna unità inserita: aggiungi una riga per ogni appartamento, con il suo prezzo e la sua data di vendita.</div>}
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {(op.vendite || []).map((v) => (
                  <div key={v.id} style={{ display: "grid", gridTemplateColumns: "1.6fr 1fr 1fr auto", gap: 10, alignItems: "end" }}>
                    <Field label="Descrizione"><input className="qg-input" style={{ fontFamily: "Inter" }} value={v.descrizione} onChange={(e) => updVendita(v.id, { descrizione: e.target.value })} placeholder="Es. Appartamento piano 1" /></Field>
                    <Field label="Prezzo (€)"><input type="number" className="qg-input" value={v.importo} onChange={(e) => updVendita(v.id, { importo: e.target.value })} /></Field>
                    <Field label="Mese vendita"><input type="month" className="qg-input" value={v.mese} onChange={(e) => updVendita(v.id, { mese: e.target.value })} /></Field>
                    <button className="qg-btn qg-btn-ghost qg-btn-danger" onClick={() => delVendita(v.id)} style={{ padding: 8 }}><Trash2 size={14} /></button>
                  </div>
                ))}
              </div>
              {(op.vendite || []).length > 0 && (
                <div style={{ display: "flex", justifyContent: "space-between", marginTop: 10, paddingTop: 10, borderTop: "1px solid var(--border-soft)", fontSize: 12.5 }}>
                  <span style={{ color: "var(--text-dim)" }}>Ricavo totale previsto</span>
                  <span className="qg-mono" style={{ fontWeight: 600 }}>{fmtEUR(ricavoTotaleOperazione(op))}</span>
                </div>
              )}
            </div>
          ) : (
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14 }}>
              <Field label="Mese vendita prevista"><input type="month" className="qg-input" value={op.meseVenditaPrevista} onChange={(e) => update({ meseVenditaPrevista: e.target.value })} /></Field>
              <Field label="Prezzo vendita previsto (€)"><input type="number" className="qg-input" value={op.prezzoVenditaPrevisto} onChange={(e) => update({ prezzoVenditaPrevisto: e.target.value })} placeholder="480000" /></Field>
              <div />
            </div>
          )}
        </div>

        {/* Costi una tantum */}
        <div className="qg-panel" style={{ padding: 20 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
            <div className="qg-label" style={{ fontSize: 12, marginBottom: 0, color: "var(--accent)" }}>Costi una tantum (acquisto, lavori, ecc.)</div>
            <button className="qg-btn qg-btn-ghost" onClick={addCostoUnaTantum}><Plus size={14} /> Aggiungi costo</button>
          </div>
          {op.costiUnaTantum.length === 0 && <div style={{ fontSize: 13, color: "var(--text-dim)" }}>Nessun costo inserito.</div>}
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {op.costiUnaTantum.map((c) => (
              <div key={c.id} style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr auto", gap: 10, alignItems: "end" }}>
                <Field label="Categoria">
                  <select className="qg-input" value={c.categoria} onChange={(e) => updCostoUnaTantum(c.id, { categoria: e.target.value })}>
                    {CAT_ONEOFF.map((cat) => <option key={cat} value={cat}>{cat}</option>)}
                  </select>
                </Field>
                <Field label="Importo (€)"><input type="number" className="qg-input" value={c.importo} onChange={(e) => updCostoUnaTantum(c.id, { importo: e.target.value })} /></Field>
                <Field label={(liquidity?.dataRiferimento && c.data && c.data <= liquidity.dataRiferimento) ? "Data · già nel saldo" : "Data"}>
                  <input type="date" className="qg-input" value={c.data} onChange={(e) => updCostoUnaTantum(c.id, { data: e.target.value })}
                    style={(liquidity?.dataRiferimento && c.data && c.data <= liquidity.dataRiferimento) ? { borderStyle: "dashed", opacity: 0.7 } : undefined}
                    title={(liquidity?.dataRiferimento && c.data && c.data <= liquidity.dataRiferimento) ? "Datata entro la riconciliazione dei saldi: si considera già pagata e non viene più proiettata. Se la spesa è slittata, sposta la data in avanti." : undefined} />
                </Field>
                <button className="qg-btn qg-btn-ghost qg-btn-danger" onClick={() => delCostoUnaTantum(c.id)} style={{ padding: 8 }}><Trash2 size={14} /></button>
              </div>
            ))}
          </div>
          {op.costiUnaTantum.length > 0 && (
            <div style={{ marginTop: 12, paddingTop: 12, borderTop: "1px solid var(--border-soft)", display: "flex", justifyContent: "flex-end", gap: 8, fontSize: 13 }}>
              <span style={{ color: "var(--text-muted)" }}>Fabbisogno iniziale totale:</span>
              <span className="qg-mono" style={{ fontWeight: 600 }}>{fmtEUR(share.fabbisognoIniziale)}</span>
            </div>
          )}
        </div>

        {/* Costi ricorrenti */}
        <div className="qg-panel" style={{ padding: 20 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
            <div className="qg-label" style={{ fontSize: 12, marginBottom: 0, color: "var(--accent)" }}>Costi ricorrenti mensili (holding cost)</div>
            <button className="qg-btn qg-btn-ghost" onClick={addCostoRicorrente}><Plus size={14} /> Aggiungi costo</button>
          </div>
          <div style={{ fontSize: 12, color: "var(--text-dim)", marginBottom: 14 }}>Questi costi restano a carico di QG per tutta la durata dell'operazione — determinano quanto a lungo l'operazione è sostenibile senza una vendita.</div>
          {op.costiRicorrenti.length === 0 && <div style={{ fontSize: 13, color: "var(--text-dim)" }}>Nessun costo ricorrente inserito.</div>}
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {op.costiRicorrenti.map((c) => (
              <div key={c.id} style={{ display: "grid", gridTemplateColumns: "1.6fr 1fr 1fr 1fr 0.8fr auto", gap: 10, alignItems: "end" }}>
                <Field label="Categoria">
                  <select className="qg-input" value={c.categoria} onChange={(e) => updCostoRicorrente(c.id, { categoria: e.target.value })}>
                    {CAT_RICORRENTE.map((cat) => <option key={cat} value={cat}>{cat}</option>)}
                  </select>
                </Field>
                <Field label="€ / mese"><input type="number" className="qg-input" value={c.importoMensile} onChange={(e) => updCostoRicorrente(c.id, { importoMensile: e.target.value })} /></Field>
                <Field label="Da"><input type="month" className="qg-input" value={c.meseInizio} onChange={(e) => updCostoRicorrente(c.id, { meseInizio: e.target.value })} /></Field>
                <Field label="A (vuoto = vendita)"><input type="month" className="qg-input" value={c.meseFine} onChange={(e) => updCostoRicorrente(c.id, { meseFine: e.target.value })} /></Field>
                <Field label="Giorno add."><input type="number" min="1" max="31" className="qg-input" value={c.giornoAddebito ?? ""} onChange={(e) => updCostoRicorrente(c.id, { giornoAddebito: e.target.value })} placeholder="fine mese" title="Giorno del mese in cui il costo viene addebitato. Serve a non riconteggiarlo se è già uscito prima della data di aggiornamento dei saldi." /></Field>
                <button className="qg-btn qg-btn-ghost qg-btn-danger" onClick={() => delCostoRicorrente(c.id)} style={{ padding: 8 }}><Trash2 size={14} /></button>
              </div>
            ))}
          </div>
        </div>

        {/* Copertura finanziaria */}
        <div className="qg-panel" style={{ padding: 20 }}>
          <div className="qg-label" style={{ fontSize: 12, marginBottom: 14, color: "var(--accent)" }}>Copertura del fabbisogno iniziale</div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
            <div className="qg-panel" style={{ padding: 14, background: "var(--panel-2)" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                <PiggyBank size={15} color="var(--accent)" />
                <span style={{ fontSize: 13, fontWeight: 600 }}>Liquidità QG</span>
              </div>
              <Field label="Importo messo da QG (€)"><input type="number" className="qg-input" value={op.copertura.liquiditaQG} onChange={(e) => updateCopertura({ liquiditaQG: e.target.value })} /></Field>
            </div>

            <div className="qg-panel" style={{ padding: 14, background: "var(--panel-2)" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <Landmark size={15} color="var(--accent)" />
                  <span style={{ fontSize: 13, fontWeight: 600 }}>{op.destinazione === "reddito" ? "Mutuo bancario" : "Prestito bancario"}</span>
                </div>
                <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: "var(--text-muted)", cursor: "pointer" }}>
                  <input type="checkbox" checked={op.copertura.prestito.attivo} onChange={(e) => updatePrestito({ attivo: e.target.checked })} /> attivo
                </label>
              </div>
              {op.copertura.prestito.attivo && (
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                    <Field label="Importo (€)"><input type="number" className="qg-input" value={op.copertura.prestito.importo} onChange={(e) => updatePrestito({ importo: e.target.value })} /></Field>
                    <Field label="Tasso annuo (%)"><input type="number" step="0.1" className="qg-input" value={op.copertura.prestito.tassoAnnuo} onChange={(e) => updatePrestito({ tassoAnnuo: e.target.value })} /></Field>
                  </div>
                  <div>
                    <Field label="Mese erogazione"><input type="month" className="qg-input" value={op.copertura.prestito.meseErogazione} onChange={(e) => updatePrestito({ meseErogazione: e.target.value })} /></Field>
                    <AvvisoCapitale ym={op.copertura.prestito.meseErogazione} />
                  </div>
                  {op.destinazione === "reddito" ? (
                    <>
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                        <Field label="Durata (mesi)"><input type="number" className="qg-input" value={op.copertura.prestito.durataMesi} onChange={(e) => updatePrestito({ durataMesi: e.target.value })} placeholder="240" /></Field>
                        <Field label="Pre-ammort. (mesi)"><input type="number" min="0" className="qg-input" value={op.copertura.prestito.mesiPreammortamento ?? ""} onChange={(e) => updatePrestito({ mesiPreammortamento: e.target.value })} placeholder="0" title="Mesi iniziali a soli interessi, compresi nella durata totale." /></Field>
                      </div>
                      <div style={{ fontSize: 12, color: "var(--text-dim)" }}>
                        {(() => {
                          const imp = Number(op.copertura.prestito.importo) || 0;
                          const durata = Math.max(0, Math.floor(Number(op.copertura.prestito.durataMesi) || 0));
                          const pre = Math.max(0, Math.min(Math.floor(Number(op.copertura.prestito.mesiPreammortamento) || 0), durata));
                          const rPre = imp * (Number(op.copertura.prestito.tassoAnnuo) || 0) / 100 / 12;
                          const rPiena = rataMutuo(imp, op.copertura.prestito.tassoAnnuo, durata - pre);
                          if (!imp || !durata) return "Inserisci importo e durata per vedere la rata.";
                          return pre > 0
                            ? <>Primi {pre} mesi (soli interessi): <b className="qg-mono">{fmtEUR(rPre)}</b>/mese · poi <b className="qg-mono">{fmtEUR(rPiena)}</b>/mese</>
                            : <>Rata mensile stimata: <b className="qg-mono">{fmtEUR(rPiena)}</b>/mese</>;
                        })()}
                      </div>
                    </>
                  ) : (
                    <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: "var(--text-muted)", cursor: "pointer" }}>
                      <input type="checkbox" checked={op.copertura.prestito.rimborsoAScadenza} onChange={(e) => updatePrestito({ rimborsoAScadenza: e.target.checked })} />
                      Rimborso a scadenza (bullet): capitale e interessi pagati alla vendita, nessuna rata mensile
                    </label>
                  )}
                </div>
              )}
            </div>

            <div className="qg-panel" style={{ padding: 14, background: "var(--panel-2)", gridColumn: "1 / -1" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <Users size={15} color="var(--accent)" />
                  <span style={{ fontSize: 13, fontWeight: 600 }}>Capitale investitori terzi</span>
                </div>
                <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: "var(--text-muted)", cursor: "pointer" }}>
                  <input type="checkbox" checked={op.copertura.investitori.attivo} onChange={(e) => updateInvestitori({ attivo: e.target.checked })} /> attivo
                </label>
              </div>
              {op.copertura.investitori.attivo && (
                <div>
                  <datalist id="qg-investor-names">
                    {(allInvestorNames || []).map((n) => <option key={n} value={n} />)}
                  </datalist>
                  {op.copertura.investitori.lista.length === 0 && (
                    <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginBottom: 10 }}>Nessun investitore inserito.</div>
                  )}
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    {op.copertura.investitori.lista.map((inv) => {
                      const quotaPct = share.fabbisognoIniziale > 0 && inv.importo ? (Number(inv.importo) / share.fabbisognoIniziale) * 100 : 0;
                      return (
                        <div key={inv.id} style={{ display: "grid", gridTemplateColumns: "1.5fr 1fr 0.9fr 0.8fr auto", gap: 10, alignItems: "end" }}>
                          <Field label="Nome investitore">
                            <input className="qg-input" style={{ fontFamily: "Inter" }} list="qg-investor-names" value={inv.nome} onChange={(e) => updInvestitore(inv.id, { nome: e.target.value })} placeholder="Es. Mario Rossi" />
                          </Field>
                          <Field label="Importo versato (€)"><input type="number" className="qg-input" value={inv.importo} onChange={(e) => updInvestitore(inv.id, { importo: e.target.value })} /></Field>
                          <div>
                            <Field label="Mese versamento"><input type="month" className="qg-input" value={inv.meseVersamento || ""} onChange={(e) => updInvestitore(inv.id, { meseVersamento: e.target.value })} title="Quando il capitale arriva sul conto QG. Serve a vedere se tra il pagamento dei costi e il versamento c'è uno scoperto." /></Field>
                            <AvvisoCapitale ym={inv.meseVersamento} />
                          </div>
                          <Field label="Quota utili">
                            <div className="qg-input qg-mono" style={{ background: "var(--panel)", color: "var(--text-muted)", display: "flex", alignItems: "center" }}>{quotaPct.toFixed(1)}%</div>
                          </Field>
                          <button className="qg-btn qg-btn-ghost qg-btn-danger" onClick={() => delInvestitore(inv.id)} style={{ padding: 8 }}><Trash2 size={14} /></button>
                        </div>
                      );
                    })}
                  </div>
                  <button type="button" className="qg-btn qg-btn-ghost" style={{ marginTop: 10, fontSize: 12 }} onClick={addInvestitore}><Plus size={14} /> Aggiungi investitore</button>
                  <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginTop: 10, lineHeight: 1.5 }}>
                    La quota di ciascun investitore è proporzionale al suo capitale versato sul fabbisogno iniziale (importo / fabbisogno × 100) e viene calcolata sull'utile netto dell'operazione (dopo IRES/IRAP), al netto della ritenuta del {liquidity?.aliquoteFiscali?.sostitutiva || "26"}% applicata prima della distribuzione. Le aliquote fiscali si impostano nella sezione Liquidità.
                  </div>
                  {share.investitori > 0 && (
                    <div style={{ fontSize: 12.5, marginTop: 10, paddingTop: 10, borderTop: "1px solid var(--border-soft)" }}>
                      <span style={{ color: "var(--text-muted)" }}>Totale capitale investitori: </span>
                      <span className="qg-mono" style={{ fontWeight: 600 }}>{fmtEUR(share.investitori)}</span>
                      <span style={{ color: "var(--text-muted)" }}> · quota complessiva: </span>
                      <span className="qg-mono" style={{ fontWeight: 600 }}>{(share.fabbisognoIniziale > 0 ? (share.investitori / share.fabbisognoIniziale) * 100 : 0).toFixed(1)}%</span>
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>

          <div style={{ marginTop: 16, paddingTop: 14, borderTop: "1px solid var(--border-soft)", display: "flex", gap: 24, fontSize: 13, flexWrap: "wrap" }}>
            <div><span style={{ color: "var(--text-muted)" }}>Fabbisogno iniziale: </span><span className="qg-mono" style={{ fontWeight: 600 }}>{fmtEUR(share.fabbisognoIniziale)}</span></div>
            <div><span style={{ color: "var(--text-muted)" }}>Copertura totale: </span><span className="qg-mono" style={{ fontWeight: 600 }}>{fmtEUR(share.coperturaTotale)}</span></div>
            <div>
              <span style={{ color: "var(--text-muted)" }}>Differenza: </span>
              <span className="qg-mono" style={{ fontWeight: 600, color: share.gap > 0 ? "var(--negative)" : "var(--positive)" }}>
                {share.gap > 0 ? `mancano ${fmtEUR(share.gap)}` : "copertura completa"}
              </span>
            </div>
          </div>
        </div>

        <div className="qg-panel" style={{ padding: 20 }}>
          <div className="qg-label" style={{ fontSize: 12, marginBottom: 12, color: "var(--accent)" }}>Anteprima economica (con imposte)</div>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
            <tbody>
              <tr style={{ borderBottom: "1px solid var(--border-soft)" }}>
                <td style={{ padding: "6px 4px", color: "var(--text-muted)" }}>Margine lordo (vendita − costi − interessi)</td>
                <td className="qg-mono" style={{ padding: "6px 4px", textAlign: "right" }}>{fmtEURSigned(econ.marginLordo)}</td>
              </tr>
              <tr style={{ borderBottom: "1px solid var(--border-soft)" }}>
                <td style={{ padding: "6px 4px", color: "var(--text-muted)" }}>IRES + IRAP ({(econ.iresPct + econ.irapPct).toFixed(1)}%)</td>
                <td className="qg-mono" style={{ padding: "6px 4px", textAlign: "right", color: "var(--negative)" }}>{fmtEURSigned(-econ.imposteSocietarie)}</td>
              </tr>
              <tr style={{ borderBottom: "1px solid var(--border-soft)" }}>
                <td style={{ padding: "6px 4px", color: "var(--text-muted)" }}>Margine netto operazione</td>
                <td className="qg-mono" style={{ padding: "6px 4px", textAlign: "right", fontWeight: 600 }}>{fmtEURSigned(econ.marginNetto)}</td>
              </tr>
              {econ.investitoriDettaglio.map((inv) => (
                <tr key={inv.id} style={{ borderBottom: "1px solid var(--border-soft)" }}>
                  <td style={{ padding: "6px 4px", color: "var(--text-muted)" }}>{inv.nome} — quota {inv.quotaPct.toFixed(1)}%, netto per lui {fmtEUR(inv.utileNetto)}</td>
                  <td className="qg-mono" style={{ padding: "6px 4px", textAlign: "right", color: "var(--negative)" }}>{fmtEURSigned(-inv.utileLordo)}</td>
                </tr>
              ))}
              <tr>
                <td style={{ padding: "8px 4px", fontWeight: 700 }}>Quota QG sull'utile</td>
                <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", fontWeight: 700, color: econ.marginQG >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(econ.marginQG)}</td>
              </tr>
            </tbody>
          </table>
          {op.copertura.investitori.attivo && econ.investitoriDettaglio.length > 0 && (
            <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginTop: 10, paddingTop: 10, borderTop: "1px dashed var(--border-soft)" }}>
              Alla vendita, oltre alla quota utili, QG restituisce a ciascun investitore anche il capitale versato: {econ.investitoriDettaglio.map((inv) => `${inv.nome} ${fmtEUR(inv.importo)}`).join(" · ")} — è un movimento di cassa che non riduce il margine, ma esce comunque dalla liquidità QG (già conteggiato nella proiezione e nell'analisi di sostenibilità).
            </div>
          )}
        </div>

        <div className="qg-panel" style={{ padding: 20 }}>
          <Field label="Note"><textarea className="qg-input" style={{ fontFamily: "Inter", minHeight: 70, resize: "vertical" }} value={op.note} onChange={(e) => update({ note: e.target.value })} placeholder="Note libere sull'operazione…" /></Field>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   OPERATION DETAIL (read-only analysis view)
   ============================================================ */
function OperationDetail({ op, liquidity, operations, rentals = [], onEdit, onBack }) {
  const analysis = useMemo(() => analyzeSustainability(op, liquidity, operations, rentals), [op, liquidity, operations, rentals]);
  const share = analysis.share;

  const opFlowsChart = useMemo(() => {
    const flows = operationQGFlows(op, liquidity, {});
    const months = Object.keys(flows).sort();
    let cum = 0;
    return months.map((m) => { cum += flows[m]; return { month: m, label: monthLabel(m), flow: flows[m], cum }; });
  }, [op, liquidity]);

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "16px 24px", borderBottom: "1px solid var(--border-soft)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button className="qg-btn qg-btn-ghost" onClick={onBack}><ChevronLeft size={16} /> Operazioni</button>
          <span className="qg-display" style={{ fontSize: 18, fontWeight: 600 }}>{op.nome}</span>
          <StatoBadge stato={op.stato} />
        </div>
        <button className="qg-btn qg-btn-primary" onClick={onEdit}>Modifica</button>
      </div>

      <div className="qg-scroll" style={{ flex: 1, overflowY: "auto", padding: 24, display: "flex", flexDirection: "column", gap: 20 }}>
        {op.indirizzo && <div style={{ fontSize: 13, color: "var(--text-muted)", marginTop: -8 }}>{op.indirizzo} · {monthLabelLong(op.meseInizio)}{analysis.aReddito ? " · a reddito" : ` → ${monthLabelLong(analysis.meseChiusura)}`}</div>}

        <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
          <VerdictBadge verdict={analysis.verdict} />
          {!analysis.profittevole && (
            <span style={{ fontSize: 13, color: "var(--negative)" }}>Margine netto previsto negativo: i costi e le imposte superano il ricavo di vendita.</span>
          )}
        </div>

        <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
          <StatCard icon={CircleDollarSign} label="Quota QG sull'utile" value={fmtCompact(analysis.marginQG)} tone={analysis.marginQG >= 0 ? "positive" : "negative"} sub={`Margine netto operazione ${fmtCompact(analysis.marginNetto)} (dopo IRES/IRAP)`} />
          <StatCard icon={Wallet} label="Fabbisogno iniziale" value={fmtCompact(share.fabbisognoIniziale)} sub={share.gap > 0 ? `mancano ${fmtCompact(share.gap)} di copertura` : "copertura completa"} tone={share.gap > 0 ? "warning" : "neutral"} />
          <StatCard icon={TrendingDown} label="Saldo minimo QG durante l'operazione" value={fmtCompact(analysis.minRow.balance)} sub={`a ${monthLabelLong(analysis.minRow.month)}, soglia min. ${fmtCompact(analysis.soglia)}`} tone={analysis.minRow.balance >= analysis.soglia ? "positive" : "negative"} />
          <StatCard icon={Clock} label="Sostenibile senza vendita per" value={analysis.mesiSostenibili === Infinity ? "> 72 mesi" : `${analysis.mesiSostenibili} mesi`} sub="tempo prima che la liquidità QG vada sotto soglia" tone={analysis.mesiSostenibili === Infinity || analysis.mesiSostenibili >= 6 ? "positive" : "warning"} />
        </div>

        {analysis.verdict === "serve_finanziamento" && (
          <div className="qg-panel" style={{ padding: 16, background: "var(--warning-soft)", border: "1px solid var(--warning)", display: "flex", gap: 10, alignItems: "flex-start" }}>
            <AlertTriangle size={16} color="var(--warning)" style={{ marginTop: 2, flexShrink: 0 }} />
            <div style={{ fontSize: 13, lineHeight: 1.5 }}>
              La liquidità propria di QG rischia di scendere sotto la soglia minima durante l'operazione (circa <b>{fmtEUR(analysis.deficitLiquidita)}</b> in meno del necessario). Per rendere l'operazione sostenibile, considera di aumentare il prestito bancario o il capitale da investitori terzi di un importo indicativo simile, riducendo la quota coperta con liquidità propria.
            </div>
          </div>
        )}
        {analysis.verdict === "sconsigliata" && (
          <div className="qg-panel" style={{ padding: 16, background: "var(--negative-soft)", border: "1px solid var(--negative)", display: "flex", gap: 10, alignItems: "flex-start" }}>
            <ShieldAlert size={16} color="var(--negative)" style={{ marginTop: 2, flexShrink: 0 }} />
            <div style={{ fontSize: 13, lineHeight: 1.5 }}>
              A prescindere da come viene finanziata, l'operazione non è in attivo: il prezzo di vendita previsto non copre costi, interessi e imposte societarie (IRES/IRAP). Rivedi il prezzo di acquisto, i costi previsti o il prezzo di vendita atteso prima di procedere.
            </div>
          </div>
        )}
        {analysis.verdict === "copertura_insufficiente" && (
          <div className="qg-panel" style={{ padding: 16, background: "var(--warning-soft)", border: "1px solid var(--warning)", display: "flex", gap: 10, alignItems: "flex-start" }}>
            <AlertTriangle size={16} color="var(--warning)" style={{ marginTop: 2, flexShrink: 0 }} />
            <div style={{ fontSize: 13, lineHeight: 1.5 }}>
              Le fonti di copertura indicate (liquidità QG + prestito + investitori) non coprono per intero il fabbisogno iniziale: mancano <b>{fmtEUR(share.gap)}</b>. Aggiungi una fonte o riduci i costi previsti.
            </div>
          </div>
        )}

        <div className="qg-panel" style={{ padding: 20 }}>
          <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 4 }}>Flusso di cassa netto QG generato dall'operazione</div>
          <div style={{ fontSize: 12, color: "var(--text-dim)", marginBottom: 10 }}>Barre = movimento del mese · linea = effetto cumulato sulla liquidità QG</div>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={opFlowsChart} margin={{ top: 6, right: 8, left: 0, bottom: 0 }}>
              <CartesianGrid stroke="var(--border-soft)" vertical={false} />
              <XAxis dataKey="label" tick={{ fill: "var(--text-dim)", fontSize: 11 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
              <YAxis tick={{ fill: "var(--text-dim)", fontSize: 11 }} axisLine={false} tickLine={false} tickFormatter={fmtCompact} width={64} />
              <Tooltip content={({ active, payload }) => {
                if (!active || !payload?.length) return null;
                const r = payload[0].payload;
                return (
                  <div className="qg-panel qg-mono" style={{ padding: "8px 10px", fontSize: 12 }}>
                    <div style={{ color: "var(--text-muted)" }}>{monthLabelLong(r.month)}</div>
                    <div style={{ color: r.flow >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(r.flow)}</div>
                    <div style={{ color: "var(--text-dim)" }}>cumulato: {fmtEURSigned(r.cum)}</div>
                  </div>
                );
              }} />
              <ReferenceLine y={0} stroke="var(--border)" />
              <Bar dataKey="flow" radius={[3, 3, 3, 3]}>
                {opFlowsChart.map((r, i) => <Cell key={i} fill={r.flow >= 0 ? "var(--positive)" : "var(--negative)"} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div className="qg-panel" style={{ padding: 20 }}>
          <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 12 }}>Riepilogo economico</div>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <tbody>
              {[
                [analysis.vendite.length > 1 ? `Ricavo totale previsto (${analysis.vendite.length} unità)` : "Prezzo di vendita previsto", analysis.ricavoTotale, "positive"],
                ["Fabbisogno iniziale (acquisto, lavori, ecc.)", -share.fabbisognoIniziale, "negative"],
                ["Costi ricorrenti totali stimati", -analysis.totaleCostiRicorrenti, "negative"],
                ["Interessi prestito bancario stimati", -analysis.interessiTotali, "negative"],
              ].map(([label, val, tone], i) => (
                <tr key={i} style={{ borderBottom: "1px solid var(--border-soft)" }}>
                  <td style={{ padding: "8px 4px", color: "var(--text-muted)" }}>{label}</td>
                  <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", color: `var(--${tone})` }}>{fmtEURSigned(val)}</td>
                </tr>
              ))}
              <tr style={{ borderBottom: "1px solid var(--border-soft)" }}>
                <td style={{ padding: "8px 4px", color: "var(--text-muted)" }}>Margine lordo</td>
                <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", fontWeight: 600 }}>{fmtEURSigned(analysis.marginLordo)}</td>
              </tr>
              <tr style={{ borderBottom: "1px solid var(--border-soft)" }}>
                <td style={{ padding: "8px 4px", color: "var(--text-muted)" }}>IRES + IRAP ({(analysis.iresPct + analysis.irapPct).toFixed(1)}%)</td>
                <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", color: "var(--negative)" }}>{fmtEURSigned(-analysis.imposteSocietarie)}</td>
              </tr>
              <tr style={{ borderBottom: "1px solid var(--border-soft)" }}>
                <td style={{ padding: "8px 4px", color: "var(--text-muted)" }}>Margine netto operazione</td>
                <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", fontWeight: 600 }}>{fmtEURSigned(analysis.marginNetto)}</td>
              </tr>
              {analysis.investitoriDettaglio.map((inv) => (
                <tr key={inv.id} style={{ borderBottom: "1px solid var(--border-soft)" }}>
                  <td style={{ padding: "8px 4px", color: "var(--text-muted)" }}>
                    {inv.nome} — quota {inv.quotaPct.toFixed(1)}% dell'utile
                    <div style={{ fontSize: 11, color: "var(--text-dim)" }}>di cui ritenuta {analysis.sostitutivaPct}%: {fmtEUR(inv.impostaSostitutiva)} · netto per lui: {fmtEUR(inv.utileNetto)}</div>
                  </td>
                  <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", color: "var(--negative)" }}>{fmtEURSigned(-inv.utileLordo)}</td>
                </tr>
              ))}
              <tr>
                <td style={{ padding: "10px 4px", fontWeight: 700 }}>Quota QG sull'utile</td>
                <td className="qg-mono" style={{ padding: "10px 4px", textAlign: "right", fontWeight: 700, color: analysis.marginQG >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(analysis.marginQG)}</td>
              </tr>
            </tbody>
          </table>
          {op.copertura.investitori.attivo && analysis.investitoriDettaglio.length > 0 && (
            <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginTop: 12, paddingTop: 12, borderTop: "1px dashed var(--border-soft)" }}>
              Alla vendita QG restituisce a ciascun investitore anche il capitale versato, in aggiunta alla quota utili. È un movimento di cassa (non un costo, non riduce il margine) già incluso nella proiezione di liquidità e nell'analisi di sostenibilità qui sopra.
            </div>
          )}
        </div>

        {op.copertura.investitori.attivo && analysis.investitoriDettaglio.length > 0 && (
          <div className="qg-panel" style={{ padding: 20 }}>
            <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 12 }}>Liquidazione investitori a fine operazione</div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ borderBottom: "1px solid var(--border-soft)" }}>
                  <th style={{ padding: "6px 4px", textAlign: "left", color: "var(--text-dim)", fontWeight: 500, fontSize: 11.5 }}>Investitore</th>
                  <th style={{ padding: "6px 4px", textAlign: "right", color: "var(--text-dim)", fontWeight: 500, fontSize: 11.5 }}>Capitale versato</th>
                  <th style={{ padding: "6px 4px", textAlign: "right", color: "var(--text-dim)", fontWeight: 500, fontSize: 11.5 }}>Quota</th>
                  <th style={{ padding: "6px 4px", textAlign: "right", color: "var(--text-dim)", fontWeight: 500, fontSize: 11.5 }}>Utile lordo</th>
                  <th style={{ padding: "6px 4px", textAlign: "right", color: "var(--text-dim)", fontWeight: 500, fontSize: 11.5 }}>Imposta 26%</th>
                  <th style={{ padding: "6px 4px", textAlign: "right", color: "var(--text-dim)", fontWeight: 500, fontSize: 11.5 }}>Utile netto</th>
                  <th style={{ padding: "6px 4px", textAlign: "right", color: "var(--text-dim)", fontWeight: 500, fontSize: 11.5 }}>Totale da liquidare</th>
                </tr>
              </thead>
              <tbody>
                {analysis.investitoriDettaglio.map((inv) => (
                  <tr key={inv.id} style={{ borderBottom: "1px solid var(--border-soft)" }}>
                    <td style={{ padding: "8px 4px" }}>{inv.nome}</td>
                    <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right" }}>{fmtEUR(inv.importo)}</td>
                    <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", color: "var(--text-muted)" }}>{inv.quotaPct.toFixed(1)}%</td>
                    <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", color: inv.utileLordo >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(inv.utileLordo)}</td>
                    <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", color: "var(--text-dim)" }}>{inv.impostaSostitutiva ? `-${fmtEUR(inv.impostaSostitutiva)}` : "—"}</td>
                    <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", color: inv.utileNetto >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(inv.utileNetto)}</td>
                    <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", fontWeight: 600 }}>{fmtEUR(inv.totaleLiquidato)}</td>
                  </tr>
                ))}
                <tr>
                  <td style={{ padding: "8px 4px", fontWeight: 700 }}>Totale</td>
                  <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", fontWeight: 700 }}>{fmtEUR(analysis.capitaleInvestitori)}</td>
                  <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", fontWeight: 700 }}>{analysis.quotaInvestitoriPct.toFixed(1)}%</td>
                  <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", fontWeight: 700 }}>{fmtEUR(analysis.utileInvestitoriLordo)}</td>
                  <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", fontWeight: 700, color: "var(--text-dim)" }}>{analysis.impostaSostitutivaInvestitori ? `-${fmtEUR(analysis.impostaSostitutivaInvestitori)}` : "—"}</td>
                  <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", fontWeight: 700 }}>{fmtEUR(analysis.utileInvestitoriNetto)}</td>
                  <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", fontWeight: 700 }}>{fmtEUR(analysis.totaleRestituitoInvestitori)}</td>
                </tr>
              </tbody>
            </table>
            <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginTop: 10 }}>
              L'utile lordo è l'importo prima della ritenuta del {analysis.sostitutivaPct}%: è la base da considerare se l'investitore lascia correre l'utile in una nuova operazione invece di prelevarlo.
            </div>
          </div>
        )}

        {op.note && (
          <div className="qg-panel" style={{ padding: 20 }}>
            <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 8 }}>Note</div>
            <div style={{ fontSize: 13, color: "var(--text-muted)", whiteSpace: "pre-wrap" }}>{op.note}</div>
          </div>
        )}
      </div>
    </div>
  );
}

/* ============================================================
   OPERATIONS LIST
   ============================================================ */
function OperationsList({ operations, liquidity, rentals = [], onOpen, onNew, filterStato, setFilterStato }) {
  const filtered = filterStato === "tutte" ? operations : operations.filter((o) => o.stato === filterStato);
  return (
    <div style={{ padding: 24 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {["tutte", ...STATI.map((s) => s.key)].map((k) => (
            <button
              key={k}
              onClick={() => setFilterStato(k)}
              className="qg-btn qg-btn-ghost"
              style={{
                fontSize: 12.5,
                background: filterStato === k ? "var(--accent-soft)" : "transparent",
                color: filterStato === k ? "var(--accent-strong)" : "var(--text-muted)",
                border: filterStato === k ? "1px solid var(--accent)" : "1px solid var(--border-soft)",
              }}
            >
              {k === "tutte" ? "Tutte" : statoInfo(k).label}
            </button>
          ))}
        </div>
        <button className="qg-btn qg-btn-primary" onClick={onNew}><Plus size={14} /> Nuova operazione</button>
      </div>

      {filtered.length === 0 ? (
        <div className="qg-panel" style={{ padding: 40, textAlign: "center", color: "var(--text-dim)" }}>
          <Building2 size={28} style={{ margin: "0 auto 12px", opacity: 0.5 }} />
          <div style={{ fontSize: 14, marginBottom: 4 }}>Nessuna operazione in questa categoria.</div>
          <div style={{ fontSize: 12.5 }}>Aggiungi la prima operazione per iniziare a monitorare costi, copertura e sostenibilità.</div>
        </div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {filtered.map((op) => {
            const analysis = analyzeSustainability(op, liquidity, operations, rentals);
            const share = analysis.share;
            return (
              <div key={op.id} onClick={() => onOpen(op.id)} className="qg-panel" style={{ padding: 16, cursor: "pointer", display: "flex", alignItems: "center", gap: 16 }}
                onMouseEnter={(e) => (e.currentTarget.style.borderColor = "var(--accent)")}
                onMouseLeave={(e) => (e.currentTarget.style.borderColor = "var(--border-soft)")}
              >
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 5 }}>
                    <span style={{ fontWeight: 600, fontSize: 14.5 }}>{op.nome || "Senza nome"}</span>
                    <StatoBadge stato={op.stato} />
                  </div>
                  <div style={{ fontSize: 12.5, color: "var(--text-muted)" }}>
                    {op.indirizzo && <>{op.indirizzo} · </>}
                    {monthLabel(op.meseInizio)} → {op.destinazione === "reddito" ? "a reddito" : monthLabel(meseChiusuraOperazione(op))}
                  </div>
                </div>
                <div className="qg-mono" style={{ textAlign: "right", minWidth: 110 }}>
                  <div style={{ fontSize: 13.5, fontWeight: 600, color: analysis.marginQG >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtCompact(analysis.marginQG)}</div>
                  <div style={{ fontSize: 11, color: "var(--text-dim)" }}>utile QG stimato</div>
                </div>
                <div className="qg-mono" style={{ textAlign: "right", minWidth: 110 }}>
                  <div style={{ fontSize: 13.5, fontWeight: 600 }}>{fmtCompact(share.liquiditaQG)}</div>
                  <div style={{ fontSize: 11, color: "var(--text-dim)" }}>liquidità QG</div>
                </div>
                <VerdictBadge verdict={analysis.verdict} />
                <ChevronRight size={16} color="var(--text-dim)" />
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ============================================================
   INVESTITORI (cross-operation live recap)
   ============================================================ */
function InvestorCard({ investor, onOpenOperation }) {
  const [open, setOpen] = useState(true);
  const gruppi = [
    { key: "pipeline", label: "Pipeline", stati: ["pipeline"] },
    { key: "in_corso", label: "In corso", stati: ["in_corso"] },
    { key: "completata", label: "Completate", stati: ["completata"] },
    { key: "abbandonata", label: "Abbandonate", stati: ["abbandonata"] },
  ];

  return (
    <div className="qg-panel" style={{ padding: 0, overflow: "hidden" }}>
      <div onClick={() => setOpen((o) => !o)} style={{ padding: 16, display: "flex", alignItems: "center", gap: 16, cursor: "pointer" }}>
        <div style={{ width: 34, height: 34, borderRadius: 8, background: "var(--accent-soft)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
          <Users size={15} color="var(--accent)" />
        </div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontWeight: 600, fontSize: 14.5 }}>{investor.nome}</div>
          <div style={{ fontSize: 12, color: "var(--text-muted)" }}>{investor.numeroOperazioni} operazion{investor.numeroOperazioni === 1 ? "e" : "i"}</div>
        </div>
        <div className="qg-mono" style={{ textAlign: "right", minWidth: 120 }}>
          <div style={{ fontSize: 13.5, fontWeight: 600 }}>{fmtCompact(investor.capitaleAttivo)}</div>
          <div style={{ fontSize: 11, color: "var(--text-dim)" }}>capitale attivo</div>
        </div>
        <div className="qg-mono" style={{ textAlign: "right", minWidth: 110 }}>
          <div style={{ fontSize: 13.5, fontWeight: 600, color: "var(--positive)" }}>{fmtCompact(investor.utileLordoProiettato)}</div>
          <div style={{ fontSize: 11, color: "var(--text-dim)" }}>utile lordo (proiez.)</div>
        </div>
        <div className="qg-mono" style={{ textAlign: "right", minWidth: 120 }}>
          <div style={{ fontSize: 13.5, fontWeight: 600, color: "var(--warning)" }}>{fmtCompact(investor.daRestituireProiettato)}</div>
          <div style={{ fontSize: 11, color: "var(--text-dim)" }}>da restituire (proiez.)</div>
        </div>
        <div className="qg-mono" style={{ textAlign: "right", minWidth: 120 }}>
          <div style={{ fontSize: 13.5, fontWeight: 600, color: "var(--positive)" }}>{fmtCompact(investor.incassatoCompletate)}</div>
          <div style={{ fontSize: 11, color: "var(--text-dim)" }}>già liquidato</div>
        </div>
        <ChevronRight size={16} color="var(--text-dim)" style={{ transform: open ? "rotate(90deg)" : "none", transition: "transform .15s" }} />
      </div>

      {open && (
        <div style={{ borderTop: "1px solid var(--border-soft)", padding: "4px 16px 16px" }}>
          {gruppi.map((g) => {
            const ops = investor.operazioni.filter((o) => g.stati.includes(o.stato));
            if (ops.length === 0) return null;
            return (
              <div key={g.key} style={{ marginTop: 14 }}>
                <div className="qg-label" style={{ fontSize: 10.5, marginBottom: 8 }}>{g.label}</div>
                <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "0 10px 5px", fontSize: 10, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.03em" }}>
                  <div style={{ flex: 1, minWidth: 0 }}>Operazione</div>
                  <div style={{ minWidth: 60, textAlign: "right" }}>Quota</div>
                  <div style={{ minWidth: 85, textAlign: "right" }}>Capitale</div>
                  <div style={{ minWidth: 85, textAlign: "right" }}>Utile lordo</div>
                  <div style={{ minWidth: 75, textAlign: "right" }}>Imposta</div>
                  <div style={{ minWidth: 85, textAlign: "right" }}>Utile netto</div>
                  <div style={{ minWidth: 95, textAlign: "right" }}>Da liquidare</div>
                  <div style={{ width: 13 }} />
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  {ops.map((o) => (
                    <div key={o.opId} onClick={() => onOpenOperation(o.opId)} style={{ display: "flex", alignItems: "center", gap: 12, padding: "8px 10px", borderRadius: 6, cursor: "pointer", background: "var(--panel-2)" }}
                      onMouseEnter={(e) => (e.currentTarget.style.background = "var(--panel-3)")}
                      onMouseLeave={(e) => (e.currentTarget.style.background = "var(--panel-2)")}
                    >
                      <div style={{ flex: 1, minWidth: 0, fontSize: 13 }}>{o.opNome}</div>
                      <div className="qg-mono" style={{ fontSize: 12, color: "var(--text-muted)", minWidth: 60, textAlign: "right" }}>{o.quotaPct.toFixed(1)}%</div>
                      <div className="qg-mono" style={{ fontSize: 12, minWidth: 85, textAlign: "right" }}>{fmtEUR(o.importo)}</div>
                      <div className="qg-mono" style={{ fontSize: 12, minWidth: 85, textAlign: "right", color: o.utileLordo >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(o.utileLordo)}</div>
                      <div className="qg-mono" style={{ fontSize: 12, minWidth: 75, textAlign: "right", color: "var(--text-dim)" }}>{o.impostaSostitutiva ? `-${fmtEUR(o.impostaSostitutiva)}` : "—"}</div>
                      <div className="qg-mono" style={{ fontSize: 12, minWidth: 85, textAlign: "right", color: o.utileNetto >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(o.utileNetto)}</div>
                      <div className="qg-mono" style={{ fontSize: 12.5, fontWeight: 600, minWidth: 95, textAlign: "right" }}>{fmtEUR(o.totaleLiquidato)}</div>
                      <ChevronRight size={13} color="var(--text-dim)" />
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function InvestorsView({ operations, liquidity, onOpenOperation }) {
  const { investors, totals } = useMemo(() => getInvestorsSummary(operations, liquidity), [operations, liquidity]);

  return (
    <div style={{ padding: 24, display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
        <StatCard icon={Users} label="Investitori coinvolti" value={String(investors.length)} sub={`su ${operations.filter((o) => o.copertura?.investitori?.attivo).length} operazioni con capitale terzi`} />
        <StatCard icon={Wallet} label="Capitale attivo (pipeline + in corso)" value={fmtCompact(totals.capitaleAttivo)} />
        <StatCard icon={Clock} label="Da restituire (proiettato)" value={fmtCompact(totals.daRestituireProiettato)} tone="warning" sub="capitale + utile netto atteso, non ancora liquidato" />
        <StatCard icon={CheckCircle2} label="Già liquidato (operazioni completate)" value={fmtCompact(totals.incassatoCompletate)} tone="positive" />
        {totals.capitaleAbbandonate > 0 && (
          <StatCard icon={ShieldAlert} label="Capitale in operazioni abbandonate" value={fmtCompact(totals.capitaleAbbandonate)} tone="negative" sub="da riconciliare manualmente" />
        )}
      </div>

      {investors.length === 0 ? (
        <div className="qg-panel" style={{ padding: 40, textAlign: "center", color: "var(--text-dim)" }}>
          <Users size={28} style={{ margin: "0 auto 12px", opacity: 0.5 }} />
          <div style={{ fontSize: 14, marginBottom: 4 }}>Nessun investitore inserito.</div>
          <div style={{ fontSize: 12.5 }}>Aggiungi investitori nella sezione "Capitale investitori terzi" di un'operazione: compariranno qui in automatico.</div>
        </div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {investors.map((inv) => (
            <InvestorCard key={inv.nome} investor={inv} onOpenOperation={onOpenOperation} />
          ))}
        </div>
      )}
    </div>
  );
}

// Explains a single month of the projection: every item that contributes to it, so any figure
// on the chart can be traced back to the record that produced it.
function DettaglioMese({ projection }) {
  const [aperto, setAperto] = useState(false);
  const [mese, setMese] = useState(currentYm());
  const riga = projection.find((r) => r.month === mese) || projection[0];
  const voci = riga?.voci || [];
  const entrate = voci.filter((v) => v.importo > 0);
  const uscite = voci.filter((v) => v.importo < 0);

  return (
    <div className="qg-panel" style={{ padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
        <div>
          <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 2, display: "flex", alignItems: "center", gap: 6 }}>
            <ListTree size={13} /> Da dove viene questo numero
          </div>
          <div style={{ fontSize: 12, color: "var(--text-dim)" }}>Scegli un mese e vedi ogni singola voce che compone entrate e uscite della proiezione.</div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <select className="qg-input" style={{ width: "auto", padding: "6px 8px" }} value={mese} onChange={(e) => { setMese(e.target.value); setAperto(true); }}>
            {projection.map((r) => <option key={r.month} value={r.month}>{monthLabelLong(r.month)}</option>)}
          </select>
          <button className="qg-btn qg-btn-ghost" onClick={() => setAperto((v) => !v)} style={{ fontSize: 11.5 }}>{aperto ? "Nascondi" : "Mostra dettaglio"}</button>
        </div>
      </div>

      {aperto && riga && (
        <div style={{ marginTop: 16 }}>
          <div style={{ display: "flex", gap: 18, flexWrap: "wrap", marginBottom: 12, fontSize: 12.5 }}>
            <span style={{ color: "var(--text-dim)" }}>Entrate <b className="qg-mono" style={{ color: "var(--positive)" }}>{fmtEUR(riga.entrate)}</b></span>
            <span style={{ color: "var(--text-dim)" }}>Uscite <b className="qg-mono" style={{ color: "var(--negative)" }}>{fmtEUR(riga.uscite)}</b></span>
            <span style={{ color: "var(--text-dim)" }}>Saldo del mese <b className="qg-mono">{fmtEURSigned(riga.net)}</b></span>
            <span style={{ color: "var(--text-dim)", marginLeft: "auto" }}>Liquidità a fine mese <b className="qg-mono">{fmtEUR(riga.balance)}</b></span>
          </div>

          {voci.length === 0 ? (
            <div style={{ fontSize: 13, color: "var(--text-dim)" }}>Nessun movimento in questo mese.</div>
          ) : (
            <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
              {[["Entrate", entrate, "var(--positive)"], ["Uscite", uscite, "var(--negative)"]].map(([titolo, lista, colore]) => (
                <div key={titolo} style={{ flex: 1, minWidth: 280 }}>
                  <div className="qg-label" style={{ marginBottom: 6, color: colore }}>{titolo}</div>
                  {lista.length === 0 ? (
                    <div style={{ fontSize: 12.5, color: "var(--text-dim)" }}>Nessuna.</div>
                  ) : (
                    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                      {lista.map((v, i) => (
                        <div key={i} style={{ display: "flex", justifyContent: "space-between", gap: 12, padding: "5px 8px", borderRadius: 5, background: "var(--panel-2)", fontSize: 12.5 }}>
                          <span style={{ minWidth: 0, overflow: "hidden", textOverflow: "ellipsis" }}>{v.voce}</span>
                          <span className="qg-mono" style={{ flexShrink: 0, color: colore }}>{fmtEUR(Math.abs(v.importo))}</span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function FlussiTooltip({ active, payload }) {
  if (!active || !payload || !payload.length) return null;
  const row = payload[0].payload;
  return (
    <div className="qg-panel qg-mono" style={{ padding: "10px 12px", fontSize: 12, border: "1px solid var(--border)" }}>
      <div style={{ color: "var(--text-muted)", marginBottom: 4 }}>{row.mese}</div>
      <div style={{ color: "var(--positive)" }}>Entrate: {fmtEUR(row.entrate)}</div>
      <div style={{ color: "var(--negative)" }}>Uscite: {fmtEUR(row.uscite)}</div>
      <div style={{ color: row.netto >= 0 ? "var(--positive)" : "var(--negative)", fontWeight: 600, marginTop: 3 }}>Saldo: {fmtEURSigned(row.netto)}</div>
    </div>
  );
}

function FlussiAnnualiChart({ liquidity, operations, rentals }) {
  const [anno, setAnno] = useState(new Date().getFullYear());
  const data = useMemo(() => computeAnnualFlows(liquidity, operations, rentals, anno), [liquidity, operations, rentals, anno]);
  const totEntrate = data.reduce((s, r) => s + r.entrate, 0);
  const totUscite = data.reduce((s, r) => s + r.uscite, 0);
  const vuoto = totEntrate === 0 && totUscite === 0;

  return (
    <div className="qg-panel" style={{ padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12, marginBottom: 14 }}>
        <div>
          <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 2 }}>Entrate e uscite previste per mese</div>
          <div style={{ fontSize: 12, color: "var(--text-dim)" }}>Movimenti pianificati dell'anno: incassi da vendite e affitti, capitale in entrata, costi, rate e imposte.</div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
          <div style={{ display: "flex", gap: 14 }}>
            <div className="qg-mono" style={{ textAlign: "right" }}>
              <div style={{ fontSize: 15, fontWeight: 600, color: "var(--positive)" }}>{fmtCompact(totEntrate)}</div>
              <div style={{ fontSize: 10.5, color: "var(--text-dim)" }}>entrate anno</div>
            </div>
            <div className="qg-mono" style={{ textAlign: "right" }}>
              <div style={{ fontSize: 15, fontWeight: 600, color: "var(--negative)" }}>{fmtCompact(totUscite)}</div>
              <div style={{ fontSize: 10.5, color: "var(--text-dim)" }}>uscite anno</div>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <button className="qg-btn qg-btn-ghost" onClick={() => setAnno((a) => a - 1)} style={{ padding: 7 }}><ChevronLeft size={15} /></button>
            <span className="qg-mono" style={{ fontSize: 14, fontWeight: 600, minWidth: 46, textAlign: "center" }}>{anno}</span>
            <button className="qg-btn qg-btn-ghost" onClick={() => setAnno((a) => a + 1)} style={{ padding: 7 }}><ChevronRight size={15} /></button>
          </div>
        </div>
      </div>

      {vuoto ? (
        <div style={{ fontSize: 13, color: "var(--text-dim)", padding: "24px 0", textAlign: "center" }}>Nessun movimento pianificato nel {anno}.</div>
      ) : (
        <>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={data} margin={{ top: 6, right: 10, left: 0, bottom: 0 }}>
              <CartesianGrid stroke="var(--border-soft)" vertical={false} />
              <XAxis dataKey="mese" tick={{ fill: "var(--text-dim)", fontSize: 11 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
              <YAxis tick={{ fill: "var(--text-dim)", fontSize: 11 }} axisLine={false} tickLine={false} tickFormatter={fmtCompact} width={60} />
              <Tooltip content={<FlussiTooltip />} cursor={{ fill: "var(--panel-2)" }} />
              <Bar dataKey="entrate" fill="var(--positive)" radius={[3, 3, 0, 0]} />
              <Bar dataKey="uscite" fill="var(--negative)" radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
          <div style={{ display: "flex", gap: 18, marginTop: 8, fontSize: 11, color: "var(--text-dim)", flexWrap: "wrap" }}>
            <span style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 10, height: 10, borderRadius: 2, background: "var(--positive)" }} /> Entrate</span>
            <span style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 10, height: 10, borderRadius: 2, background: "var(--negative)" }} /> Uscite</span>
            <span style={{ marginLeft: "auto" }}>Sono flussi <b>pianificati</b>: i mesi già passati mostrano quanto era previsto, non il consuntivo.</span>
          </div>
        </>
      )}
    </div>
  );
}

/* ============================================================
   SCADENZARIO SPESE (day-level one-off expenses)
   ============================================================ */
function SpeseSchedule({ operations, liquidity, onOpen }) {
  const [mostraPassate, setMostraPassate] = useState(false);
  const oggi = todayStr();
  const cutoff = liquidity?.dataRiferimento || null;

  const tutte = useMemo(() => buildSpeseSchedule(operations, liquidity), [operations, liquidity]);
  const future = tutte.filter((s) => s.data >= oggi);
  const passate = tutte.filter((s) => s.data < oggi);
  const visibili = mostraPassate ? tutte : future;

  // Group by date for a clean, dated timeline.
  const gruppi = useMemo(() => {
    const map = new Map();
    visibili.forEach((s) => {
      if (!map.has(s.data)) map.set(s.data, []);
      map.get(s.data).push(s);
    });
    return Array.from(map.entries()).sort((a, b) => (a[0] < b[0] ? -1 : 1));
  }, [visibili]);

  const totaleFuturo = future.reduce((s, x) => s + x.importo, 0);
  const prossimi30 = future.filter((s) => monthDiffDays(oggi, s.data) <= 30).reduce((s, x) => s + x.importo, 0);

  return (
    <div className="qg-panel" style={{ padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12, marginBottom: 14 }}>
        <div>
          <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 2, display: "flex", alignItems: "center", gap: 6 }}>
            <CalendarClock size={13} /> Scadenzario spese
          </div>
          <div style={{ fontSize: 12, color: "var(--text-dim)" }}>Spese una tantum di operazioni e costi fissi, ordinate per giorno. I costi ricorrenti mensili non compaiono qui.</div>
        </div>
        <div style={{ display: "flex", gap: 18, alignItems: "center" }}>
          <div className="qg-mono" style={{ textAlign: "right" }}>
            <div style={{ fontSize: 17, fontWeight: 600 }}>{fmtCompact(prossimi30)}</div>
            <div style={{ fontSize: 10.5, color: "var(--text-dim)" }}>prossimi 30 giorni</div>
          </div>
          <div className="qg-mono" style={{ textAlign: "right" }}>
            <div style={{ fontSize: 17, fontWeight: 600 }}>{fmtCompact(totaleFuturo)}</div>
            <div style={{ fontSize: 10.5, color: "var(--text-dim)" }}>totale in programma</div>
          </div>
        </div>
      </div>

      {passate.length > 0 && (
        <button
          className="qg-btn qg-btn-ghost"
          style={{ fontSize: 11.5, marginBottom: 12, padding: "5px 10px" }}
          onClick={() => setMostraPassate((v) => !v)}
        >
          {mostraPassate ? "Nascondi spese passate" : `Mostra anche ${passate.length} spes${passate.length === 1 ? "a passata" : "e passate"}`}
        </button>
      )}

      {gruppi.length === 0 ? (
        <div style={{ fontSize: 13, color: "var(--text-dim)" }}>Nessuna spesa una tantum in programma.</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          {gruppi.map(([data, righe]) => {
            const passata = data < oggi;
            const giaNelSaldo = cutoff && data <= cutoff;
            const totaleGiorno = righe.reduce((s, x) => s + x.importo, 0);
            const giorniAllaScadenza = monthDiffDays(oggi, data);
            return (
              <div key={data} style={{ opacity: passata ? 0.55 : 1 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                    <span style={{ width: 8, height: 8, borderRadius: "50%", background: passata ? "var(--text-dim)" : giorniAllaScadenza <= 30 ? "var(--warning)" : "var(--accent)", flexShrink: 0 }} />
                    <span style={{ fontSize: 13, fontWeight: 600 }}>{dateLabel(data)}</span>
                    {!passata && (
                      <span style={{ fontSize: 11, color: "var(--text-dim)" }}>
                        {giorniAllaScadenza === 0 ? "oggi" : giorniAllaScadenza === 1 ? "domani" : `fra ${giorniAllaScadenza} giorni`}
                      </span>
                    )}
                    {giaNelSaldo && (
                      <span style={{ fontSize: 10, padding: "2px 7px", borderRadius: 5, background: "var(--panel-2)", color: "var(--text-dim)", border: "1px solid var(--border-soft)" }} title="Datata entro la riconciliazione dei saldi: si considera già pagata e non viene più proiettata. Se è slittata, aggiorna la data della spesa.">già nel saldo</span>
                    )}
                  </div>
                  <span className="qg-mono" style={{ fontSize: 13, fontWeight: 600, color: "var(--negative)" }}>{fmtEURSigned(-totaleGiorno)}</span>
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 4, paddingLeft: 18, borderLeft: "1px solid var(--border-soft)", marginLeft: 3 }}>
                  {righe.map((s) => (
                    <div
                      key={s.id}
                      onClick={s.opId ? () => onOpen(s.opId) : undefined}
                      style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, padding: "5px 8px", borderRadius: 6, cursor: s.opId ? "pointer" : "default" }}
                      onMouseEnter={(e) => { if (s.opId) e.currentTarget.style.background = "var(--panel-2)"; }}
                      onMouseLeave={(e) => { e.currentTarget.style.background = "transparent"; }}
                    >
                      <div style={{ minWidth: 0 }}>
                        <div style={{ fontSize: 12.5, color: "var(--text)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{s.categoria}</div>
                        <div style={{ fontSize: 11, color: "var(--text-dim)", display: "flex", alignItems: "center", gap: 5 }}>
                          {s.tipo === "fisso" ? <Receipt size={10} /> : <Building2 size={10} />}
                          {s.origine}
                        </div>
                      </div>
                      <span className="qg-mono" style={{ fontSize: 12.5, color: "var(--text-muted)", flexShrink: 0 }}>{fmtEUR(s.importo)}</span>
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ============================================================
   DASHBOARD
   ============================================================ */
function Dashboard({ liquidity, operations, rentals = [], onOpen, onOpenRental }) {
  const [horizon, setHorizon] = useState(24);
  const [orizzonteCard, setOrizzonteCard] = useState(12); // 3 / 6 / 12 mesi nella card di sintesi
  const projection = useMemo(() => computeGlobalProjection(liquidity, operations, rentals, horizon), [liquidity, operations, rentals, horizon]);
  const soglia = Number(liquidity.soglioMinimo) || 0;

  const attive = operations.filter((o) => o.stato === "in_corso" || o.stato === "pipeline");
  const totaleAllocato = attive.reduce((s, o) => s + fundingShare(o).liquiditaQG, 0);
  const saldoAttuale = totaleDisponibile(liquidity);
  const vincolato = totaleVincolato(liquidity);
  const lordo = totaleLiquidita(liquidity);
  const saldoAOrizzonte = projection[Math.min(orizzonteCard - 1, projection.length - 1)]?.balance ?? saldoAttuale;
  const capitaleInvestitori = capitaleInvestitoriAttivo(operations);
  const debitoBancario = finanziamentiAttivi(operations, rentals);
  const minProiettato = projection.reduce((min, r) => (r.balance < min.balance ? r : min), projection[0] || { balance: saldoAttuale, month: currentYm() });
  const critiche = attive.filter((o) => analyzeSustainability(o, liquidity, operations, rentals).verdict !== "ok");
  const now = currentYm();
  const costiFissiMensili = (liquidity.costiFissi?.ricorrenti || []).reduce((sum, c) => {
    const attivo = c.meseInizio && monthDiff(c.meseInizio, now) >= 0 && (!c.meseFine || monthDiff(now, c.meseFine) >= 0);
    return attivo ? sum + (Number(c.importoMensile) || 0) : sum;
  }, 0);

  // Cash flow mensile netto degli immobili a reddito attivi (affitto − costi − rata mutuo).
  const rentalsAttivi = (rentals || []).filter((r) => r.attivo);
  const cashFlowRentaleMensile = rentalsAttivi.reduce((s, r) => s + rentalEconomics(r, opDiRental(r, operations)).cashFlowAnnuoNetto / 12, 0);

  // Portafoglio titoli: riserva liquidabile, non cassa operativa. Mostrata come cuscinetto.
  const port = portfolioSummary(liquidity.portafoglio, liquidity.fx?.usdToEur);
  const valoreTitoli = port.valoreAttuale;
  const projectionConTitoli = useMemo(
    () => projection.map((r) => ({ ...r, balanceConTitoli: r.balance + valoreTitoli })),
    [projection, valoreTitoli]
  );

  return (
    <div style={{ padding: 24, display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
        <StatCard icon={Wallet} label={vincolato > 0 ? "Liquidità disponibile (netta)" : "Liquidità disponibile ora"} value={fmtCompact(saldoAttuale)} sub={vincolato > 0
          ? `${fmtCompact(lordo)} in banca, di cui ${fmtCompact(vincolato)} di terzi`
          : ((liquidity.conti && liquidity.conti.length > 1) ? `su ${liquidity.conti.length} conti${liquidity.dataRiferimento ? ` · al ${dateLabel(liquidity.dataRiferimento)}` : ""}` : (liquidity.dataRiferimento ? `al ${dateLabel(liquidity.dataRiferimento)}` : undefined))} />
        <StatCard
          icon={TrendingUp}
          label={`Liquidità prevista (${orizzonteCard}m)`}
          value={fmtCompact(saldoAOrizzonte)}
          tone={saldoAOrizzonte >= soglia ? "positive" : "negative"}
          sub={`a ${monthLabelLong(addMonths(currentYm(), orizzonteCard - 1))}`}
          extra={
            <span style={{ display: "flex", gap: 3 }}>
              {[3, 6, 12].map((m) => (
                <button
                  key={m}
                  onClick={() => setOrizzonteCard(m)}
                  className="qg-btn qg-btn-ghost"
                  style={{
                    fontSize: 10.5, padding: "2px 6px", lineHeight: 1.3,
                    background: orizzonteCard === m ? "var(--accent-soft)" : "transparent",
                    color: orizzonteCard === m ? "var(--accent-strong)" : "var(--text-dim)",
                    border: orizzonteCard === m ? "1px solid var(--accent)" : "1px solid var(--border-soft)",
                  }}
                >{m}m</button>
              ))}
            </span>
          }
        />
        <StatCard icon={TrendingDown} label={`Minimo proiettato (${horizon} mesi)`} value={fmtCompact(minProiettato.balance)} sub={`a ${monthLabelLong(minProiettato.month)}`} tone={minProiettato.balance >= soglia ? "positive" : "negative"} />
        <StatCard icon={Building2} label="Capitale QG allocato in operazioni" value={fmtCompact(totaleAllocato)} sub={`su ${attive.length} operazion${attive.length === 1 ? "e" : "i"} attive`} />
        <StatCard icon={Users} label="Capitale investitori gestito" value={fmtCompact(capitaleInvestitori)} tone="warning" sub="capitale di terzi in operazioni aperte, da restituire" />
        <StatCard icon={Landmark} label="Finanziamenti bancari attivi" value={fmtCompact(debitoBancario)} tone="warning" sub="debito residuo verso le banche" />
        {rentalsAttivi.length > 0 && <StatCard icon={KeyRound} label="Rendita netta mensile da affitti" value={fmtCompact(cashFlowRentaleMensile)} sub={`da ${rentalsAttivi.length} immobil${rentalsAttivi.length === 1 ? "e" : "i"} a reddito`} tone={cashFlowRentaleMensile >= 0 ? "positive" : "negative"} />}
        {valoreTitoli > 0 && <StatCard icon={LineChartIcon} label="Riserva titoli (liquidabile)" value={fmtCompact(valoreTitoli)} sub={`P/L ${fmtEURSigned(port.pl)}${port.investito > 0 ? ` · ${fmtPct(port.plPct)}` : ""}`} tone={port.pl >= 0 ? "positive" : "neutral"} />}
        <StatCard icon={Receipt} label="Costi fissi mensili correnti" value={fmtCompact(costiFissiMensili)} sub="stipendi, ufficio e altri costi aziendali" />
      </div>

      {critiche.length > 0 && (
        <div className="qg-panel" style={{ padding: 16, background: "var(--warning-soft)", border: "1px solid var(--warning)", display: "flex", gap: 10, alignItems: "flex-start" }}>
          <AlertTriangle size={16} color="var(--warning)" style={{ marginTop: 2, flexShrink: 0 }} />
          <div style={{ fontSize: 13 }}>
            <b>{critiche.length}</b> operazion{critiche.length === 1 ? "e richiede" : "i richiedono"} attenzione: {critiche.map((o) => o.nome).join(", ")}.
          </div>
        </div>
      )}

      <div className="qg-panel" style={{ padding: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <div>
            <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 2 }}>Proiezione liquidità QG</div>
            <div style={{ fontSize: 12, color: "var(--text-dim)" }}>Saldo attuale + movimenti + costi fissi + debiti/crediti + flussi di operazioni e immobili a reddito{valoreTitoli > 0 ? ". La linea tratteggiata aggiunge la riserva titoli, liquidabile ma non cassa operativa." : ""}</div>
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            {[12, 24, 36, 60].map((h) => (
              <button key={h} onClick={() => setHorizon(h)} className="qg-btn qg-btn-ghost" style={{
                fontSize: 12, padding: "6px 10px",
                background: horizon === h ? "var(--accent-soft)" : "transparent",
                color: horizon === h ? "var(--accent-strong)" : "var(--text-muted)",
                border: horizon === h ? "1px solid var(--accent)" : "1px solid var(--border-soft)",
              }}>{h}m</button>
            ))}
          </div>
        </div>
        <ResponsiveContainer width="100%" height={280}>
          <AreaChart data={projectionConTitoli} margin={{ top: 6, right: 10, left: 0, bottom: 0 }}>
            <defs>
              <linearGradient id="qgLiquidityFill" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor="var(--accent)" stopOpacity={0.35} />
                <stop offset="100%" stopColor="var(--accent)" stopOpacity={0} />
              </linearGradient>
            </defs>
            <CartesianGrid stroke="var(--border-soft)" vertical={false} />
            <XAxis dataKey="label" tick={{ fill: "var(--text-dim)", fontSize: 11 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} interval={Math.ceil(horizon / 12)} />
            <YAxis tick={{ fill: "var(--text-dim)", fontSize: 11 }} axisLine={false} tickLine={false} tickFormatter={fmtCompact} width={64} />
            <Tooltip content={<LiquidityTooltip valoreTitoli={valoreTitoli} />} />
            {soglia !== 0 && <ReferenceLine y={soglia} stroke="var(--warning)" strokeDasharray="4 4" label={{ value: "soglia minima", fill: "var(--warning)", fontSize: 10, position: "insideTopLeft" }} />}
            <ReferenceLine y={0} stroke="var(--border)" />
            {valoreTitoli > 0 && <Area type="monotone" dataKey="balanceConTitoli" stroke="var(--text-dim)" strokeWidth={1.5} strokeDasharray="5 4" fill="none" />}
            <Area type="monotone" dataKey="balance" stroke="var(--accent)" strokeWidth={2} fill="url(#qgLiquidityFill)" />
          </AreaChart>
        </ResponsiveContainer>
        {valoreTitoli > 0 && (
          <div style={{ display: "flex", gap: 18, marginTop: 10, fontSize: 11, color: "var(--text-dim)" }}>
            <span style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 14, height: 2, background: "var(--accent)", display: "inline-block" }} /> Liquidità operativa</span>
            <span style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 14, height: 0, borderTop: "2px dashed var(--text-dim)", display: "inline-block" }} /> Liquidità + riserva titoli</span>
          </div>
        )}
      </div>

      <DettaglioMese projection={projection} />

      <FlussiAnnualiChart liquidity={liquidity} operations={operations} rentals={rentals} />

      <SpeseSchedule operations={operations} liquidity={liquidity} onOpen={onOpen} />

      <div className="qg-panel" style={{ padding: 20 }}>
        <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 14 }}>Allocazione capitale per operazione</div>
        {attive.length === 0 ? (
          <div style={{ fontSize: 13, color: "var(--text-dim)" }}>Nessuna operazione attiva o in pipeline al momento.</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {attive.map((op) => {
              const s = fundingShare(op);
              const pct = totaleAllocato > 0 ? (s.liquiditaQG / totaleAllocato) * 100 : 0;
              const analysis = analyzeSustainability(op, liquidity, operations);
              return (
                <div key={op.id} onClick={() => onOpen(op.id)} style={{ cursor: "pointer" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 4 }}>
                    <span style={{ display: "flex", alignItems: "center", gap: 8 }}>{op.nome} <StatoBadge stato={op.stato} /></span>
                    <span className="qg-mono">{fmtEUR(s.liquiditaQG)}</span>
                  </div>
                  <div style={{ height: 6, background: "var(--panel-2)", borderRadius: 4, overflow: "hidden" }}>
                    <div style={{ height: "100%", width: `${pct}%`, background: analysis.verdict === "ok" ? "var(--accent)" : "var(--warning)", borderRadius: 4 }} />
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

/* ============================================================
   COSTI FISSI (company overhead: salaries, office, etc.)
   ============================================================ */
function CostiFissiSettings({ liquidity, setLiquidity }) {
  const costiFissi = liquidity.costiFissi || { unaTantum: [], ricorrenti: [] };
  const updateCostiFissi = (patch) => setLiquidity((prev) => ({ ...prev, costiFissi: { ...(prev.costiFissi || { unaTantum: [], ricorrenti: [] }), ...patch } }));

  const addUnaTantum = () => updateCostiFissi({ unaTantum: [...costiFissi.unaTantum, { id: uid(), categoria: CAT_FISSI_UNA_TANTUM[0], importo: "", data: todayStr(), descrizione: "" }] });
  const updUnaTantum = (id, patch) => updateCostiFissi({ unaTantum: costiFissi.unaTantum.map((c) => (c.id === id ? { ...c, ...patch } : c)) });
  const delUnaTantum = (id) => updateCostiFissi({ unaTantum: costiFissi.unaTantum.filter((c) => c.id !== id) });

  const addRicorrente = () => updateCostiFissi({ ricorrenti: [...costiFissi.ricorrenti, { id: uid(), categoria: CAT_FISSI_RICORRENTE[0], importoMensile: "", meseInizio: currentYm(), meseFine: "", descrizione: "" }] });
  const updRicorrente = (id, patch) => updateCostiFissi({ ricorrenti: costiFissi.ricorrenti.map((c) => (c.id === id ? { ...c, ...patch } : c)) });
  const delRicorrente = (id) => updateCostiFissi({ ricorrenti: costiFissi.ricorrenti.filter((c) => c.id !== id) });

  const now = currentYm();
  const totaleMensileAttuale = costiFissi.ricorrenti.reduce((sum, c) => {
    const attivo = c.meseInizio && monthDiff(c.meseInizio, now) >= 0 && (!c.meseFine || monthDiff(now, c.meseFine) >= 0);
    return attivo ? sum + (Number(c.importoMensile) || 0) : sum;
  }, 0);

  return (
    <div style={{ padding: 24, display: "flex", flexDirection: "column", gap: 20, maxWidth: 920 }}>
      <div className="qg-panel" style={{ padding: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
          <div>
            <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 4 }}>Costi fissi aziendali</div>
            <div style={{ fontSize: 12.5, color: "var(--text-dim)", maxWidth: 480 }}>Stipendi, ufficio e altri costi di struttura, indipendenti dalle singole operazioni immobiliari. Vengono inclusi automaticamente nella proiezione di liquidità QG.</div>
          </div>
          <div className="qg-mono" style={{ textAlign: "right" }}>
            <div style={{ fontSize: 22, fontWeight: 600 }}>{fmtEUR(totaleMensileAttuale)}</div>
            <div style={{ fontSize: 11, color: "var(--text-dim)" }}>burn rate mensile attuale</div>
          </div>
        </div>
      </div>

      <div className="qg-panel" style={{ padding: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
          <div className="qg-label" style={{ fontSize: 12, marginBottom: 0, color: "var(--accent)" }}>Costi ricorrenti mensili</div>
          <button className="qg-btn qg-btn-ghost" onClick={addRicorrente}><Plus size={14} /> Aggiungi costo</button>
        </div>
        <div style={{ fontSize: 12, color: "var(--text-dim)", marginBottom: 14 }}>Es. stipendi, affitto sede, consulenze fisse. Lascia vuoto "A" se il costo prosegue senza una data di fine prevista. Il <b>giorno di addebito</b> è opzionale: indicandolo, il costo del mese in corso non viene riconteggiato se è già uscito prima della data di aggiornamento dei saldi (vuoto = trattato come fine mese).</div>
        {costiFissi.ricorrenti.length === 0 && <div style={{ fontSize: 13, color: "var(--text-dim)" }}>Nessun costo ricorrente inserito.</div>}
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {costiFissi.ricorrenti.map((c) => (
            <div key={c.id} style={{ display: "grid", gridTemplateColumns: "1.3fr 0.9fr 0.9fr 0.9fr 0.7fr 1.3fr auto", gap: 10, alignItems: "end" }}>
              <Field label="Categoria">
                <select className="qg-input" value={c.categoria} onChange={(e) => updRicorrente(c.id, { categoria: e.target.value })}>
                  {CAT_FISSI_RICORRENTE.map((cat) => <option key={cat} value={cat}>{cat}</option>)}
                </select>
              </Field>
              <Field label="€ / mese"><input type="number" className="qg-input" value={c.importoMensile} onChange={(e) => updRicorrente(c.id, { importoMensile: e.target.value })} /></Field>
              <Field label="Da"><input type="month" className="qg-input" value={c.meseInizio} onChange={(e) => updRicorrente(c.id, { meseInizio: e.target.value })} /></Field>
              <Field label="A (opzionale)"><input type="month" className="qg-input" value={c.meseFine} onChange={(e) => updRicorrente(c.id, { meseFine: e.target.value })} /></Field>
              <Field label="Giorno add."><input type="number" min="1" max="31" className="qg-input" value={c.giornoAddebito ?? ""} onChange={(e) => updRicorrente(c.id, { giornoAddebito: e.target.value })} placeholder="fine mese" title="Giorno del mese in cui il costo viene addebitato. Serve a non riconteggiarlo se è già uscito prima della data di aggiornamento dei saldi." /></Field>
              <Field label="Descrizione"><input className="qg-input" style={{ fontFamily: "Inter" }} value={c.descrizione} onChange={(e) => updRicorrente(c.id, { descrizione: e.target.value })} placeholder="Es. Stipendio Mario Rossi" /></Field>
              <button className="qg-btn qg-btn-ghost qg-btn-danger" onClick={() => delRicorrente(c.id)} style={{ padding: 8 }}><Trash2 size={14} /></button>
            </div>
          ))}
        </div>
      </div>

      <div className="qg-panel" style={{ padding: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
          <div className="qg-label" style={{ fontSize: 12, marginBottom: 0, color: "var(--accent)" }}>Costi una tantum</div>
          <button className="qg-btn qg-btn-ghost" onClick={addUnaTantum}><Plus size={14} /> Aggiungi costo</button>
        </div>
        <div style={{ fontSize: 12, color: "var(--text-dim)", marginBottom: 14 }}>Es. attrezzature, consulenze straordinarie, imposte societarie non legate a una specifica operazione.</div>
        {costiFissi.unaTantum.length === 0 && <div style={{ fontSize: 13, color: "var(--text-dim)" }}>Nessun costo una tantum inserito.</div>}
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {costiFissi.unaTantum.map((c) => (
            <div key={c.id} style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr 1fr 1.6fr auto", gap: 10, alignItems: "end" }}>
              <Field label="Categoria">
                <select className="qg-input" value={c.categoria} onChange={(e) => updUnaTantum(c.id, { categoria: e.target.value })}>
                  {CAT_FISSI_UNA_TANTUM.map((cat) => <option key={cat} value={cat}>{cat}</option>)}
                </select>
              </Field>
              <Field label="Importo (€)"><input type="number" className="qg-input" value={c.importo} onChange={(e) => updUnaTantum(c.id, { importo: e.target.value })} /></Field>
              <Field label={(liquidity?.dataRiferimento && c.data && c.data <= liquidity.dataRiferimento) ? "Data · già nel saldo" : "Data"}>
                <input type="date" className="qg-input" value={c.data} onChange={(e) => updUnaTantum(c.id, { data: e.target.value })}
                  style={(liquidity?.dataRiferimento && c.data && c.data <= liquidity.dataRiferimento) ? { borderStyle: "dashed", opacity: 0.7 } : undefined}
                  title={(liquidity?.dataRiferimento && c.data && c.data <= liquidity.dataRiferimento) ? "Datata entro la riconciliazione dei saldi: si considera già pagata e non viene più proiettata. Se la spesa è slittata, sposta la data in avanti." : undefined} />
              </Field>
              <Field label="Descrizione"><input className="qg-input" style={{ fontFamily: "Inter" }} value={c.descrizione} onChange={(e) => updUnaTantum(c.id, { descrizione: e.target.value })} placeholder="Es. Nuovi arredi ufficio" /></Field>
              <button className="qg-btn qg-btn-ghost qg-btn-danger" onClick={() => delUnaTantum(c.id)} style={{ padding: 8 }}><Trash2 size={14} /></button>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   LIQUIDITY SETTINGS
   ============================================================ */
function LiquiditySettings({ liquidity, setLiquidity }) {
  const addMovimento = () => setLiquidity((prev) => ({ ...prev, movimenti: [...prev.movimenti, { id: uid(), data: todayStr(), importo: "", descrizione: "" }] }));
  const updMovimento = (id, patch) => setLiquidity((prev) => ({ ...prev, movimenti: prev.movimenti.map((m) => (m.id === id ? { ...m, ...patch } : m)) }));
  const delMovimento = (id) => setLiquidity((prev) => ({ ...prev, movimenti: prev.movimenti.filter((m) => m.id !== id) }));
  const updateAliquote = (patch) => setLiquidity((prev) => ({ ...prev, aliquoteFiscali: { ...prev.aliquoteFiscali, ...patch } }));

  const partite = liquidity.partite || [];
  const addPartita = (tipo) => setLiquidity((prev) => ({ ...prev, partite: [...(prev.partite || []), { id: uid(), tipo, controparte: "", importo: "", data: todayStr(), saldato: false, descrizione: "" }] }));
  const updPartita = (id, patch) => setLiquidity((prev) => ({ ...prev, partite: (prev.partite || []).map((p) => (p.id === id ? { ...p, ...patch } : p)) }));
  const delPartita = (id) => setLiquidity((prev) => ({ ...prev, partite: (prev.partite || []).filter((p) => p.id !== id) }));

  const oggi = todayStr();
  const crediti = partite.filter((p) => p.tipo === "credito");
  const debiti = partite.filter((p) => p.tipo === "debito");
  const totCreditiAperti = crediti.filter((p) => !p.saldato).reduce((s, p) => s + (Number(p.importo) || 0), 0);
  const totDebitiAperti = debiti.filter((p) => !p.saldato).reduce((s, p) => s + (Number(p.importo) || 0), 0);

  const titoli = liquidity.portafoglio?.titoli || [];
  const addTitolo = () => setLiquidity((prev) => ({ ...prev, portafoglio: { titoli: [...(prev.portafoglio?.titoli || []), { id: uid(), nome: "", ticker: "", quantita: "", prezzoCarico: "", prezzoAttuale: "", valuta: "EUR" }] } }));
  const updTitolo = (id, patch) => setLiquidity((prev) => ({ ...prev, portafoglio: { titoli: (prev.portafoglio?.titoli || []).map((t) => (t.id === id ? { ...t, ...patch } : t)) } }));
  const delTitolo = (id) => setLiquidity((prev) => ({ ...prev, portafoglio: { titoli: (prev.portafoglio?.titoli || []).filter((t) => t.id !== id) } }));

  const fx = liquidity.fx || { usdToEur: "", aggiornato: "" };
  const port = portfolioSummary(liquidity.portafoglio, fx.usdToEur);
  const [fxLoading, setFxLoading] = useState(false);
  const [fxError, setFxError] = useState("");
  const refreshFx = async () => {
    setFxLoading(true); setFxError("");
    try {
      const { usdToEur, date } = await fetchFxRate();
      setLiquidity((prev) => ({ ...prev, fx: { usdToEur: String(usdToEur), aggiornato: date } }));
    } catch (e) {
      setFxError("Cambio non recuperabile ora — puoi inserirlo a mano.");
    } finally {
      setFxLoading(false);
    }
  };

  // --- Aggiornamento quotazioni (via Claude + ricerca web), con anteprima prima di applicare ---
  const [quoteLoading, setQuoteLoading] = useState(false);
  const [quoteError, setQuoteError] = useState("");
  const [quotePreview, setQuotePreview] = useState(null); // [{ id, nome, ticker, vecchio, nuovo, valutaVecchia, valutaNuova, data, fonte, ok }]

  const aggiornaQuotazioni = async () => {
    setQuoteLoading(true); setQuoteError(""); setQuotePreview(null);
    try {
      const res = await fetchQuotesViaClaude(titoli);
      if (res.error) { setQuoteError(res.error); return; }
      const byTicker = {};
      res.quotes.forEach((q) => { if (q && q.ticker) byTicker[String(q.ticker).trim().toUpperCase()] = q; });
      const righe = titoli.map((t) => {
        const key = (t.ticker || "").trim().toUpperCase();
        const q = key ? byTicker[key] : null;
        const nuovo = q && q.prezzo != null && Number.isFinite(Number(q.prezzo)) ? Number(q.prezzo) : null;
        const valutaNuova = q && q.valuta ? String(q.valuta).toUpperCase() : null;
        return {
          id: t.id, nome: t.nome || t.ticker || "titolo", ticker: t.ticker || "",
          vecchio: Number(t.prezzoAttuale) || 0,
          nuovo,
          valutaVecchia: t.valuta || "EUR",
          valutaNuova: valutaNuova && ["EUR", "USD"].includes(valutaNuova) ? valutaNuova : null,
          data: q?.data || "", fonte: q?.fonte || "",
          ok: nuovo != null,
        };
      });
      if (!righe.some((r) => r.ok)) { setQuoteError("Non ho trovato quotazioni affidabili per i ticker inseriti. Verifica che i simboli siano corretti."); return; }
      setQuotePreview(righe);
    } catch (e) {
      setQuoteError("Non sono riuscito a recuperare le quotazioni. Riprova fra un momento.");
    } finally {
      setQuoteLoading(false);
    }
  };

  const applicaQuotazioni = () => {
    if (!quotePreview) return;
    setLiquidity((prev) => ({
      ...prev,
      portafoglio: {
        titoli: (prev.portafoglio?.titoli || []).map((t) => {
          const r = quotePreview.find((x) => x.id === t.id && x.ok);
          if (!r) return t;
          const patch = { prezzoAttuale: String(r.nuovo) };
          if (r.valutaNuova && r.valutaNuova !== t.valuta) patch.valuta = r.valutaNuova;
          return { ...t, ...patch };
        }),
      },
    }));
    setQuotePreview(null);
  };

  const fileRef = useRef(null);
  const [imp, setImp] = useState(null); // { movimenti, saldo, saldoData, meta } | { error }

  const conti = liquidity.conti && liquidity.conti.length ? liquidity.conti : [{ id: "legacy", nome: "Conto principale", saldo: liquidity.saldoAttuale ?? "" }];
  const addConto = () => setLiquidity((prev) => ({ ...prev, conti: [...(prev.conti || []), { id: uid(), nome: `Conto ${(prev.conti?.length || 0) + 1}`, saldo: "", vincolato: "" }] }));
  const updConto = (id, patch) => setLiquidity((prev) => ({ ...prev, conti: (prev.conti || []).map((c) => (c.id === id ? { ...c, ...patch } : c)) }));
  const delConto = (id) => setLiquidity((prev) => ({ ...prev, conti: (prev.conti || []).filter((c) => c.id !== id) }));
  const totaleContiLordo = conti.reduce((s, c) => s + (Number(c.saldo) || 0), 0);
  const totaleContiVincolato = conti.reduce((s, c) => s + (Number(c.vincolato) || 0), 0);

  const [contoImportId, setContoImportId] = useState(null);
  const contoImportEffettivo = contoImportId || conti[0]?.id;

  const handleFile = async (file) => {
    if (!file) return;
    setImp({ loading: true });
    try {
      const res = await parseBankFile(file);
      setImp(res);
    } catch (e) {
      setImp({ error: "Impossibile leggere il file." });
    }
  };

  // Signature to detect duplicates when re-importing an overlapping statement.
  const movSig = (m) => `${m.data}|${Math.round(Number(m.importo) * 100)}|${(m.descrizione || "").trim().toLowerCase()}`;

  const confirmImport = (opts = {}) => {
    if (!imp || imp.error) return;
    setLiquidity((prev) => {
      const esistenti = new Set((prev.movimenti || []).map(movSig));
      const nuovi = imp.movimenti.filter((m) => !esistenti.has(movSig(m)));
      const next = { ...prev, movimenti: [...(prev.movimenti || []), ...nuovi] };
      if (opts.aggiornaSaldo && imp.saldo != null) {
        const targetId = contoImportEffettivo;
        next.conti = (prev.conti || []).map((c) => (c.id === targetId ? { ...c, saldo: String(imp.saldo) } : c));
        if (imp.saldoData) next.dataRiferimento = imp.saldoData;
      }
      return next;
    });
    setImp(null);
    if (fileRef.current) fileRef.current.value = "";
  };

  const cancelImport = () => {
    setImp(null);
    if (fileRef.current) fileRef.current.value = "";
  };

  const nuoviConteggio = useMemo(() => {
    if (!imp || imp.error || imp.loading) return 0;
    const esistenti = new Set((liquidity.movimenti || []).map(movSig));
    return imp.movimenti.filter((m) => !esistenti.has(movSig(m))).length;
  }, [imp, liquidity.movimenti]);

  return (
    <div style={{ padding: 24, display: "flex", flexDirection: "column", gap: 20, maxWidth: 780 }}>
      <div className="qg-panel" style={{ padding: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 14 }}>
          <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 0 }}>Liquidità disponibile</div>
          <button className="qg-btn qg-btn-ghost" onClick={addConto}><Plus size={14} /> Aggiungi conto</button>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1.6fr 1fr 1fr auto", gap: 10, padding: "0 4px 6px", fontSize: 10.5, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.04em" }}>
          <div>Conto</div><div style={{ textAlign: "right" }}>Saldo (€)</div><div style={{ textAlign: "right" }}>Di cui di terzi (€)</div><div />
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {conti.map((c) => (
            <div key={c.id} style={{ display: "grid", gridTemplateColumns: "1.6fr 1fr 1fr auto", gap: 10, alignItems: "center" }}>
              <input className="qg-input" style={{ fontFamily: "Inter" }} value={c.nome} onChange={(e) => updConto(c.id, { nome: e.target.value })} placeholder="Es. Unicredit business" />
              <input type="number" className="qg-input" style={{ textAlign: "right" }} value={c.saldo} onChange={(e) => updConto(c.id, { saldo: e.target.value })} placeholder="0" />
              <input type="number" className="qg-input" style={{ textAlign: "right" }} value={c.vincolato ?? ""} onChange={(e) => updConto(c.id, { vincolato: e.target.value })} placeholder="0" title="Denaro presente sul conto ma di proprietà di terzi (es. quota dei proprietari sugli incassi Smart Rent). Viene escluso dalla liquidità spendibile." />
              <button className="qg-btn qg-btn-ghost qg-btn-danger" onClick={() => delConto(c.id)} style={{ padding: 8 }} disabled={conti.length <= 1} title={conti.length <= 1 ? "Deve restare almeno un conto" : "Elimina conto"}><Trash2 size={14} /></button>
            </div>
          ))}
        </div>
        <div style={{ marginTop: 12, paddingTop: 12, borderTop: "1px solid var(--border-soft)", display: "flex", flexDirection: "column", gap: 5 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5 }}>
            <span style={{ color: "var(--text-dim)" }}>Totale in banca</span>
            <span className="qg-mono">{fmtEUR(totaleContiLordo)}</span>
          </div>
          {totaleContiVincolato > 0 && (
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5 }}>
              <span style={{ color: "var(--text-dim)" }}>− Denaro di terzi (non spendibile)</span>
              <span className="qg-mono" style={{ color: "var(--warning)" }}>{fmtEUR(totaleContiVincolato)}</span>
            </div>
          )}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", paddingTop: 5, borderTop: totaleContiVincolato > 0 ? "1px solid var(--border-soft)" : "none" }}>
            <span style={{ fontSize: 12.5, fontWeight: 600 }}>Liquidità disponibile</span>
            <span className="qg-mono" style={{ fontSize: 16, fontWeight: 600 }}>{fmtEUR(totaleContiLordo - totaleContiVincolato)}</span>
          </div>
        </div>
        <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginTop: 10 }}>
          La colonna "di cui di terzi" serve per conti su cui transita denaro non vostro — ad esempio la quota dei proprietari sugli incassi gestiti da Smart Rent. Inserisci il saldo pieno (quello che vedi in banca) e a fianco quanto è dovuto a terzi: la proiezione userà solo la parte disponibile.
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginTop: 16 }}>
          <Field label="Saldi aggiornati al"><input type="date" className="qg-input" value={liquidity.dataRiferimento} onChange={(e) => setLiquidity((prev) => ({ ...prev, dataRiferimento: e.target.value }))} /></Field>
          <Field label="Soglia minima di sicurezza (€)"><input type="number" className="qg-input" value={liquidity.soglioMinimo} onChange={(e) => setLiquidity((prev) => ({ ...prev, soglioMinimo: e.target.value }))} placeholder="0" /></Field>
        </div>
        <div style={{ fontSize: 12, color: "var(--text-dim)", marginTop: 10 }}>La proiezione parte dalla somma dei conti. Tutto ciò che è datato <b>fino alla data di aggiornamento inclusa</b> si considera già dentro questi saldi e non viene proiettato di nuovo (niente doppi conteggi). La soglia minima è il livello di cassa complessivo sotto il quale un'operazione viene segnalata come rischiosa.</div>
      </div>

      {/* Import estratto conto (CSV / Excel) */}
      <div className="qg-panel" style={{ padding: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12, marginBottom: 12 }}>
          <div>
            <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 4, display: "flex", alignItems: "center", gap: 6 }}>
              <FileUp size={13} /> Importa estratto conto Unicredit (CSV o Excel)
            </div>
            <div style={{ fontSize: 12.5, color: "var(--text-dim)", maxWidth: 520, lineHeight: 1.5 }}>
              Scarica l'estratto conto da Unicredit in formato Excel (.xls / .xlsx) o CSV e caricalo qui: leggo automaticamente data, importo e descrizione di ogni movimento. Il collegamento diretto e automatico al conto non è possibile in questa app — serve un aggiornamento manuale del file.
            </div>
          </div>
          <input ref={fileRef} type="file" accept=".csv,.xls,.xlsx,.xlsm,.txt,text/csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" style={{ display: "none" }} onChange={(e) => handleFile(e.target.files?.[0])} />
          <button className="qg-btn qg-btn-primary" onClick={() => fileRef.current?.click()}><Upload size={14} /> Carica file</button>
        </div>

        {imp && imp.loading && (
          <div style={{ fontSize: 12.5, color: "var(--text-muted)" }}>Lettura del file in corso…</div>
        )}

        {imp && imp.error && (
          <div className="qg-panel" style={{ padding: 12, background: "var(--negative-soft)", border: "1px solid var(--negative)", fontSize: 12.5, display: "flex", gap: 8, alignItems: "center" }}>
            <ShieldAlert size={14} color="var(--negative)" style={{ flexShrink: 0 }} />
            <span>{imp.error} Se il tuo export usa un formato particolare, dimmelo e adatto il lettore.</span>
          </div>
        )}

        {imp && !imp.error && !imp.loading && (
          <div className="qg-panel" style={{ padding: 16, background: "var(--panel-2)", border: "1px solid var(--border)" }}>
            <div style={{ fontSize: 13, marginBottom: 10 }}>
              Ho letto <b>{imp.movimenti.length}</b> moviment{imp.movimenti.length === 1 ? "o" : "i"} dal file{imp.saldo != null && <> · saldo finale rilevato <b className="qg-mono">{fmtEUR(imp.saldo)}</b>{imp.saldoData && <> al {dateLabel(imp.saldoData)}</>}</>}.
              {nuoviConteggio < imp.movimenti.length && (
                <div style={{ fontSize: 12, color: "var(--text-dim)", marginTop: 4 }}>{imp.movimenti.length - nuoviConteggio} già presenti (verranno saltati), {nuoviConteggio} nuovi da aggiungere.</div>
              )}
            </div>

            <div className="qg-scroll" style={{ maxHeight: 180, overflowY: "auto", marginBottom: 12, border: "1px solid var(--border-soft)", borderRadius: 6 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ position: "sticky", top: 0, background: "var(--panel-3)" }}>
                    <th style={{ padding: "6px 8px", textAlign: "left", color: "var(--text-dim)", fontWeight: 500 }}>Data</th>
                    <th style={{ padding: "6px 8px", textAlign: "left", color: "var(--text-dim)", fontWeight: 500 }}>Descrizione</th>
                    <th style={{ padding: "6px 8px", textAlign: "right", color: "var(--text-dim)", fontWeight: 500 }}>Importo</th>
                  </tr>
                </thead>
                <tbody>
                  {imp.movimenti.slice(0, 60).map((m) => (
                    <tr key={m.id} style={{ borderTop: "1px solid var(--border-soft)" }}>
                      <td className="qg-mono" style={{ padding: "5px 8px", whiteSpace: "nowrap" }}>{dateLabel(m.data)}</td>
                      <td style={{ padding: "5px 8px", color: "var(--text-muted)", maxWidth: 280, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{m.descrizione || "—"}</td>
                      <td className="qg-mono" style={{ padding: "5px 8px", textAlign: "right", color: Number(m.importo) >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(Number(m.importo))}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {imp.movimenti.length > 60 && <div style={{ padding: "6px 8px", fontSize: 11, color: "var(--text-dim)" }}>…e altri {imp.movimenti.length - 60} movimenti</div>}
            </div>

            {imp.saldo != null && conti.length > 1 && (
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, fontSize: 12.5 }}>
                <span style={{ color: "var(--text-dim)" }}>Aggiorna il saldo del conto:</span>
                <select className="qg-input" style={{ width: "auto", padding: "6px 8px" }} value={contoImportEffettivo} onChange={(e) => setContoImportId(e.target.value)}>
                  {conti.map((c) => <option key={c.id} value={c.id}>{c.nome}</option>)}
                </select>
              </div>
            )}

            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              {imp.saldo != null && (
                <button className="qg-btn qg-btn-primary" onClick={() => confirmImport({ aggiornaSaldo: true })}>
                  <CheckCircle2 size={14} /> Aggiungi movimenti e aggiorna saldo{conti.length > 1 ? ` di "${conti.find((c) => c.id === contoImportEffettivo)?.nome || ""}"` : ""}
                </button>
              )}
              <button className="qg-btn" onClick={() => confirmImport({ aggiornaSaldo: false })}>
                Aggiungi solo i movimenti
              </button>
              <button className="qg-btn qg-btn-ghost" onClick={cancelImport}>Annulla</button>
            </div>
          </div>
        )}
      </div>

      <div className="qg-panel" style={{ padding: 20 }}>
        <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 14 }}>Aliquote fiscali</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14 }}>
          <Field label="IRES (%)"><input type="number" step="0.1" className="qg-input" value={liquidity.aliquoteFiscali?.ires ?? ""} onChange={(e) => updateAliquote({ ires: e.target.value })} /></Field>
          <Field label="IRAP (%)"><input type="number" step="0.1" className="qg-input" value={liquidity.aliquoteFiscali?.irap ?? ""} onChange={(e) => updateAliquote({ irap: e.target.value })} /></Field>
          <Field label="Imposta sostitutiva su distribuzione investitori (%)"><input type="number" step="0.1" className="qg-input" value={liquidity.aliquoteFiscali?.sostitutiva ?? ""} onChange={(e) => updateAliquote({ sostitutiva: e.target.value })} /></Field>
        </div>
        <div style={{ fontSize: 12, color: "var(--text-dim)", marginTop: 10 }}>
          Applicate a ogni operazione: IRES e IRAP riducono il margine lordo prima del riparto; la sostitutiva si applica solo sulla quota di utile distribuita agli investitori terzi, prima che la ricevano.
        </div>
      </div>

      <div className="qg-panel" style={{ padding: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
          <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 0 }}>Movimenti extra (non legati a un'operazione)</div>
          <button className="qg-btn qg-btn-ghost" onClick={addMovimento}><Plus size={14} /> Aggiungi movimento</button>
        </div>
        <div style={{ fontSize: 12, color: "var(--text-dim)", marginBottom: 14 }}>Entrate o uscite <b>future</b> non modellate altrove: apporti di capitale, F24, rimborsi attesi. I costi del mese già passato non vanno qui — sono già nel saldo che hai aggiornato.</div>
        {liquidity.movimenti.length === 0 && <div style={{ fontSize: 13, color: "var(--text-dim)" }}>Nessun movimento extra inserito.</div>}
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {liquidity.movimenti.slice().sort((a, b) => (a.data < b.data ? -1 : 1)).map((m) => {
            const giaNelSaldo = liquidity.dataRiferimento && m.data && m.data <= liquidity.dataRiferimento;
            return (
              <div key={m.id} style={{ display: "grid", gridTemplateColumns: "1fr 1fr 2fr auto", gap: 10, alignItems: "end", opacity: giaNelSaldo ? 0.5 : 1 }}>
                <Field label={giaNelSaldo ? "Data · già nel saldo" : "Data"}>
                  <input type="date" className="qg-input" value={m.data} onChange={(e) => updMovimento(m.id, { data: e.target.value })} style={giaNelSaldo ? { borderStyle: "dashed" } : undefined} title={giaNelSaldo ? "Antecedente alla data di aggiornamento: già incluso nei saldi, non viene proiettato. Se il pagamento è slittato, sposta la data in avanti." : undefined} />
                </Field>
                <Field label="Importo (€, +/-)"><input type="number" className="qg-input" value={m.importo} onChange={(e) => updMovimento(m.id, { importo: e.target.value })} /></Field>
                <Field label="Descrizione"><input className="qg-input" style={{ fontFamily: "Inter" }} value={m.descrizione} onChange={(e) => updMovimento(m.id, { descrizione: e.target.value })} placeholder="Es. apporto soci" /></Field>
                <button className="qg-btn qg-btn-ghost qg-btn-danger" onClick={() => delMovimento(m.id)} style={{ padding: 8 }}><Trash2 size={14} /></button>
              </div>
            );
          })}
        </div>
        {liquidity.movimenti.some((m) => liquidity.dataRiferimento && m.data && m.data <= liquidity.dataRiferimento) && (
          <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginTop: 12, display: "flex", gap: 7, alignItems: "flex-start" }}>
            <CheckCircle2 size={13} color="var(--text-dim)" style={{ flexShrink: 0, marginTop: 1 }} />
            <span>Le righe in grigio sono datate fino al {dateLabel(liquidity.dataRiferimento)}: si considerano già incluse nei saldi dei conti, quindi non vengono proiettate una seconda volta. <b>Se uno di questi pagamenti è slittato</b>, sposta la sua data in avanti: tornerà automaticamente nella proiezione.</span>
          </div>
        )}
      </div>

      {/* Debiti e crediti con scadenza */}
      <div className="qg-panel" style={{ padding: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12, marginBottom: 6 }}>
          <div>
            <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 4, display: "flex", alignItems: "center", gap: 6 }}>
              <CircleDollarSign size={13} /> Debiti e crediti con scadenza
            </div>
            <div style={{ fontSize: 12.5, color: "var(--text-dim)", maxWidth: 560, lineHeight: 1.5 }}>
              Crediti da incassare ed debiti da rimborsare, ciascuno con la sua data. Entrano nella proiezione di liquidità alla data indicata, finché non li segni come saldati.
            </div>
          </div>
          <div style={{ display: "flex", gap: 18 }}>
            <div className="qg-mono" style={{ textAlign: "right" }}>
              <div style={{ fontSize: 17, fontWeight: 600, color: "var(--positive)" }}>{fmtCompact(totCreditiAperti)}</div>
              <div style={{ fontSize: 10.5, color: "var(--text-dim)" }}>crediti da incassare</div>
            </div>
            <div className="qg-mono" style={{ textAlign: "right" }}>
              <div style={{ fontSize: 17, fontWeight: 600, color: "var(--negative)" }}>{fmtCompact(totDebitiAperti)}</div>
              <div style={{ fontSize: 10.5, color: "var(--text-dim)" }}>debiti da rimborsare</div>
            </div>
          </div>
        </div>

        <div style={{ display: "flex", gap: 8, margin: "12px 0 14px" }}>
          <button className="qg-btn qg-btn-ghost" onClick={() => addPartita("credito")}><ArrowDownLeft size={14} /> Aggiungi credito</button>
          <button className="qg-btn qg-btn-ghost" onClick={() => addPartita("debito")}><ArrowUpRight size={14} /> Aggiungi debito</button>
        </div>

        {partite.length === 0 ? (
          <div style={{ fontSize: 13, color: "var(--text-dim)" }}>Nessun debito o credito inserito.</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {partite.slice().sort((a, b) => (a.data < b.data ? -1 : 1)).map((p) => {
              const scaduta = !p.saldato && p.data < oggi;
              const isCredito = p.tipo === "credito";
              return (
                <div key={p.id} style={{ display: "grid", gridTemplateColumns: "auto 1.4fr 1fr 1fr 1.6fr auto auto", gap: 10, alignItems: "end", opacity: p.saldato ? 0.55 : 1 }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "center", width: 30, height: 34, borderRadius: 7, background: isCredito ? "var(--positive-soft)" : "var(--negative-soft)", flexShrink: 0 }} title={isCredito ? "Credito (in entrata)" : "Debito (in uscita)"}>
                    {isCredito ? <ArrowDownLeft size={14} color="var(--positive)" /> : <ArrowUpRight size={14} color="var(--negative)" />}
                  </div>
                  <Field label="Controparte"><input className="qg-input" style={{ fontFamily: "Inter" }} value={p.controparte} onChange={(e) => updPartita(p.id, { controparte: e.target.value })} placeholder={isCredito ? "Chi deve pagare" : "A chi devi pagare"} /></Field>
                  <Field label="Importo (€)"><input type="number" className="qg-input" value={p.importo} onChange={(e) => updPartita(p.id, { importo: e.target.value })} /></Field>
                  <Field label={isCredito ? "Da incassare il" : "Da rimborsare il"}>
                    <input type="date" className="qg-input" value={p.data} onChange={(e) => updPartita(p.id, { data: e.target.value })} style={scaduta ? { borderColor: "var(--warning)" } : undefined} />
                  </Field>
                  <Field label="Descrizione"><input className="qg-input" style={{ fontFamily: "Inter" }} value={p.descrizione} onChange={(e) => updPartita(p.id, { descrizione: e.target.value })} placeholder="Note (facoltativo)" /></Field>
                  <Field label={p.saldato ? "Saldato" : "Aperto"}>
                    <button
                      className="qg-btn qg-btn-ghost"
                      onClick={() => updPartita(p.id, { saldato: !p.saldato })}
                      style={{ padding: "8px 10px", fontSize: 11.5, color: p.saldato ? "var(--positive)" : "var(--text-muted)", whiteSpace: "nowrap" }}
                      title="Segna come saldato/aperto"
                    >
                      {p.saldato ? <><CheckCircle2 size={13} /> Sì</> : "Segna"}
                    </button>
                  </Field>
                  <button className="qg-btn qg-btn-ghost qg-btn-danger" onClick={() => delPartita(p.id)} style={{ padding: 8 }}><Trash2 size={14} /></button>
                </div>
              );
            })}
          </div>
        )}
        <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginTop: 12 }}>
          Nota: una partita "saldata" si considera già regolata e riflessa nel saldo attuale, quindi non viene più proiettata. Una scadenza già passata e non saldata resta nella proiezione ed è evidenziata.
        </div>
      </div>

      {/* Portafoglio titoli */}
      <div className="qg-panel" style={{ padding: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12, marginBottom: 6 }}>
          <div>
            <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 4, display: "flex", alignItems: "center", gap: 6 }}>
              <LineChartIcon size={13} /> Portafoglio titoli
            </div>
            <div style={{ fontSize: 12.5, color: "var(--text-dim)", maxWidth: 560, lineHeight: 1.5 }}>
              Azioni e titoli detenuti. Sono una riserva liquidabile ma non cassa operativa: non entrano nella proiezione di liquidità, ma li vedi in Dashboard come cuscinetto disponibile. Con "Aggiorna quotazioni" cerco i prezzi sul web e te li mostro in anteprima; puoi sempre correggerli a mano.
            </div>
          </div>
          <div style={{ display: "flex", gap: 18 }}>
            <div className="qg-mono" style={{ textAlign: "right" }}>
              <div style={{ fontSize: 17, fontWeight: 600 }}>{fmtCompact(port.valoreAttuale)}</div>
              <div style={{ fontSize: 10.5, color: "var(--text-dim)" }}>valore attuale</div>
            </div>
            <div className="qg-mono" style={{ textAlign: "right" }}>
              <div style={{ fontSize: 17, fontWeight: 600, color: port.pl >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(port.pl)}</div>
              <div style={{ fontSize: 10.5, color: "var(--text-dim)" }}>P/L {port.investito > 0 ? `(${fmtPct(port.plPct)})` : ""}</div>
            </div>
          </div>
        </div>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, margin: "14px 0" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", fontSize: 12 }}>
            <span style={{ color: "var(--text-dim)" }}>Cambio USD → EUR:</span>
            <input
              type="number" step="0.0001"
              className="qg-input qg-mono"
              style={{ width: 100, padding: "6px 8px" }}
              value={fx.usdToEur}
              onChange={(e) => setLiquidity((prev) => ({ ...prev, fx: { usdToEur: e.target.value, aggiornato: "manuale" } }))}
              placeholder="0,92"
            />
            <button className="qg-btn qg-btn-ghost" style={{ fontSize: 11.5, padding: "6px 10px" }} onClick={refreshFx} disabled={fxLoading}>
              {fxLoading ? "Aggiorno…" : "Aggiorna cambio"}
            </button>
            {fx.aggiornato && fx.aggiornato !== "manuale" && <span style={{ color: "var(--text-dim)", fontSize: 11 }}>BCE, {dateLabel(fx.aggiornato)}</span>}
            {fx.aggiornato === "manuale" && <span style={{ color: "var(--text-dim)", fontSize: 11 }}>inserito a mano</span>}
            {fxError && <span style={{ color: "var(--warning)", fontSize: 11 }}>{fxError}</span>}
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button className="qg-btn qg-btn-ghost" onClick={aggiornaQuotazioni} disabled={quoteLoading || titoli.length === 0} title="Cerca sul web l'ultima quotazione dei ticker inseriti e te la mostra in anteprima prima di applicarla">
              <RefreshCw size={14} /> {quoteLoading ? "Cerco quotazioni…" : "Aggiorna quotazioni"}
            </button>
            <button className="qg-btn qg-btn-ghost" onClick={addTitolo}><Plus size={14} /> Aggiungi titolo</button>
          </div>
        </div>

        {quoteError && (
          <div className="qg-panel" style={{ padding: 11, background: "var(--warning-soft)", border: "1px solid var(--warning)", fontSize: 12.5, marginBottom: 12, display: "flex", gap: 8, alignItems: "center" }}>
            <AlertTriangle size={14} color="var(--warning)" style={{ flexShrink: 0 }} /> {quoteError}
          </div>
        )}

        {quotePreview && (
          <div className="qg-panel" style={{ padding: 16, background: "var(--panel-2)", border: "1px solid var(--border)", marginBottom: 14 }}>
            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 4 }}>Quotazioni trovate — controlla prima di applicare</div>
            <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginBottom: 12 }}>Prezzi reperiti dal web: possono avere ritardo o riferirsi alla chiusura precedente. Verifica prima di confermare.</div>
            <div className="qg-scroll" style={{ maxHeight: 240, overflowY: "auto", border: "1px solid var(--border-soft)", borderRadius: 6, marginBottom: 12 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ position: "sticky", top: 0, background: "var(--panel-3)" }}>
                    <th style={{ padding: "6px 8px", textAlign: "left", color: "var(--text-dim)", fontWeight: 500 }}>Titolo</th>
                    <th style={{ padding: "6px 8px", textAlign: "right", color: "var(--text-dim)", fontWeight: 500 }}>Attuale</th>
                    <th style={{ padding: "6px 8px", textAlign: "right", color: "var(--text-dim)", fontWeight: 500 }}>Nuovo</th>
                    <th style={{ padding: "6px 8px", textAlign: "right", color: "var(--text-dim)", fontWeight: 500 }}>Var.</th>
                    <th style={{ padding: "6px 8px", textAlign: "left", color: "var(--text-dim)", fontWeight: 500 }}>Fonte / data</th>
                  </tr>
                </thead>
                <tbody>
                  {quotePreview.map((r) => {
                    const varPct = r.ok && r.vecchio > 0 ? ((r.nuovo - r.vecchio) / r.vecchio) * 100 : null;
                    const cambioValuta = r.ok && r.valutaNuova && r.valutaNuova !== r.valutaVecchia;
                    return (
                      <tr key={r.id} style={{ borderTop: "1px solid var(--border-soft)", opacity: r.ok ? 1 : 0.5 }}>
                        <td style={{ padding: "5px 8px" }}>
                          {r.nome}{r.ticker && <span className="qg-mono" style={{ color: "var(--text-dim)", fontSize: 11 }}> · {r.ticker.toUpperCase()}</span>}
                          {cambioValuta && <div style={{ fontSize: 10.5, color: "var(--warning)" }}>valuta aggiornata: {r.valutaVecchia} → {r.valutaNuova}</div>}
                        </td>
                        <td className="qg-mono" style={{ padding: "5px 8px", textAlign: "right", color: "var(--text-dim)" }}>{r.vecchio ? r.vecchio.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : "—"}</td>
                        <td className="qg-mono" style={{ padding: "5px 8px", textAlign: "right", fontWeight: 600 }}>{r.ok ? r.nuovo.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : "non trovato"}</td>
                        <td className="qg-mono" style={{ padding: "5px 8px", textAlign: "right", color: varPct == null ? "var(--text-dim)" : varPct >= 0 ? "var(--positive)" : "var(--negative)" }}>{varPct == null ? "—" : `${varPct >= 0 ? "+" : ""}${fmtPct(varPct)}`}</td>
                        <td style={{ padding: "5px 8px", color: "var(--text-dim)", fontSize: 11 }}>{[r.fonte, r.data].filter(Boolean).join(" · ") || "—"}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <button className="qg-btn qg-btn-primary" onClick={applicaQuotazioni}><CheckCircle2 size={14} /> Applica le quotazioni trovate</button>
              <button className="qg-btn qg-btn-ghost" onClick={() => setQuotePreview(null)}>Annulla</button>
            </div>
          </div>
        )}

        {port.hasUsd && !(Number(fx.usdToEur) > 0) && (
          <div className="qg-panel" style={{ padding: 10, background: "var(--warning-soft)", border: "1px solid var(--warning)", fontSize: 12, marginBottom: 12, display: "flex", gap: 8, alignItems: "center" }}>
            <AlertTriangle size={13} color="var(--warning)" style={{ flexShrink: 0 }} />
            <span>Hai titoli in dollari ma manca il cambio: i valori USD non sono convertiti. Premi "Aggiorna cambio" o inseriscilo a mano.</span>
          </div>
        )}

        {titoli.length === 0 ? (
          <div style={{ fontSize: 13, color: "var(--text-dim)" }}>Nessun titolo inserito.</div>
        ) : (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "1.5fr 0.7fr 0.7fr 0.8fr 0.9fr 0.9fr 1fr auto", gap: 8, padding: "0 4px 6px", fontSize: 10, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.03em" }}>
              <div>Titolo</div><div>Ticker</div><div>Valuta</div><div style={{ textAlign: "right" }}>Quantità</div><div style={{ textAlign: "right" }}>Pr. carico</div><div style={{ textAlign: "right" }}>Pr. attuale</div><div style={{ textAlign: "right" }}>Valore / P.L (€)</div><div />
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {port.titoli.map((t) => {
                const simbolo = t.valuta === "USD" ? "$" : "€";
                return (
                  <div key={t.id} style={{ display: "grid", gridTemplateColumns: "1.5fr 0.7fr 0.7fr 0.8fr 0.9fr 0.9fr 1fr auto", gap: 8, alignItems: "center" }}>
                    <input className="qg-input" style={{ fontFamily: "Inter" }} value={t.nome} onChange={(e) => updTitolo(t.id, { nome: e.target.value })} placeholder="Es. Apple Inc." />
                    <input className="qg-input" style={{ fontFamily: "IBM Plex Mono", textTransform: "uppercase" }} value={t.ticker} onChange={(e) => updTitolo(t.id, { ticker: e.target.value })} placeholder="AAPL" />
                    <select className="qg-input" value={t.valuta} onChange={(e) => updTitolo(t.id, { valuta: e.target.value })}>
                      <option value="EUR">€ EUR</option>
                      <option value="USD">$ USD</option>
                    </select>
                    <input type="number" className="qg-input" style={{ textAlign: "right" }} value={t.quantita} onChange={(e) => updTitolo(t.id, { quantita: e.target.value })} placeholder="0" />
                    <input type="number" step="0.01" className="qg-input" style={{ textAlign: "right" }} value={t.prezzoCarico} onChange={(e) => updTitolo(t.id, { prezzoCarico: e.target.value })} placeholder={`0,00 ${simbolo}`} />
                    <input type="number" step="0.01" className="qg-input" style={{ textAlign: "right" }} value={t.prezzoAttuale} onChange={(e) => updTitolo(t.id, { prezzoAttuale: e.target.value })} placeholder={`0,00 ${simbolo}`} />
                    <div className="qg-mono" style={{ textAlign: "right", fontSize: 12 }}>
                      <div style={{ fontWeight: 600 }}>{fmtEUR(t.valoreAttuale)}</div>
                      <div style={{ fontSize: 10.5, color: t.pl >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(t.pl)} · {fmtPct(t.plPct)}</div>
                    </div>
                    <button className="qg-btn qg-btn-ghost qg-btn-danger" onClick={() => delTitolo(t.id)} style={{ padding: 8 }}><Trash2 size={14} /></button>
                  </div>
                );
              })}
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 14, paddingTop: 12, borderTop: "1px solid var(--border-soft)", fontSize: 12.5, flexWrap: "wrap", gap: 8 }}>
              <span style={{ color: "var(--text-dim)" }}>Investito: <b className="qg-mono">{fmtEUR(port.investito)}</b></span>
              <span style={{ color: "var(--text-dim)" }}>Valore attuale: <b className="qg-mono" style={{ color: "var(--text)" }}>{fmtEUR(port.valoreAttuale)}</b> · P/L complessivo: <b className="qg-mono" style={{ color: port.pl >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(port.pl)} ({fmtPct(port.plPct)})</b></span>
            </div>
            {port.hasUsd && Number(fx.usdToEur) > 0 && <div style={{ fontSize: 11, color: "var(--text-dim)", marginTop: 8 }}>Prezzi mostrati nella valuta del titolo; valori e P/L convertiti in euro al cambio 1 $ = {Number(fx.usdToEur).toLocaleString("it-IT", { minimumFractionDigits: 4, maximumFractionDigits: 4 })} €.</div>}
          </>
        )}
      </div>
    </div>
  );
}

/* ============================================================
   IMMOBILI A REDDITO (buy-to-let / short-let)
   ============================================================ */
function RentalMetric({ label, value, sub, tone = "neutral", mono = true }) {
  const color = tone === "positive" ? "var(--positive)" : tone === "negative" ? "var(--negative)" : tone === "accent" ? "var(--accent)" : "var(--text)";
  return (
    <div className="qg-panel" style={{ padding: "14px 16px", flex: 1, minWidth: 150 }}>
      <div className="qg-label" style={{ marginBottom: 6 }}>{label}</div>
      <div className={mono ? "qg-mono" : ""} style={{ fontSize: 20, fontWeight: 600, color, lineHeight: 1.1 }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: "var(--text-dim)", marginTop: 4 }}>{sub}</div>}
    </div>
  );
}

function RentalsList({ rentals, operations = [], onOpen, onNew }) {
  const attivi = rentals.filter((r) => r.attivo);
  const cashFlowTotMensile = attivi.reduce((s, r) => s + rentalEconomics(r, opDiRental(r, operations)).cashFlowAnnuoNetto / 12, 0);
  const investimentoTot = rentals.reduce((s, r) => s + rentalEconomics(r, opDiRental(r, operations)).investimentoTotale, 0);

  return (
    <div style={{ padding: 24, display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
        <div>
          <div className="qg-display" style={{ fontSize: 20, fontWeight: 600 }}>Immobili a reddito</div>
          <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginTop: 2 }}>Immobili tenuti in pancia e affittati (idealmente affitti brevi), con costi, ricavi stimati e rendimento.</div>
        </div>
        <button className="qg-btn qg-btn-primary" onClick={onNew}><Plus size={15} /> Nuovo immobile</button>
      </div>

      {rentals.length > 0 && (
        <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
          <RentalMetric label="Rendita netta mensile" value={fmtCompact(cashFlowTotMensile)} tone={cashFlowTotMensile >= 0 ? "positive" : "negative"} sub={`${attivi.length} immobil${attivi.length === 1 ? "e" : "i"} in locazione`} />
          <RentalMetric label="Capitale investito" value={fmtCompact(investimentoTot)} sub={`${rentals.length} immobil${rentals.length === 1 ? "e" : "i"} totali`} />
        </div>
      )}

      {rentals.length === 0 ? (
        <div className="qg-panel" style={{ padding: 40, textAlign: "center" }}>
          <KeyRound size={26} color="var(--text-dim)" style={{ marginBottom: 10 }} />
          <div style={{ fontSize: 14, marginBottom: 4 }}>Nessun immobile a reddito</div>
          <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginBottom: 16 }}>Aggiungi un immobile destinato all'affitto per stimarne rendimento e sostenibilità.</div>
          <button className="qg-btn qg-btn-primary" onClick={onNew} style={{ margin: "0 auto" }}><Plus size={15} /> Nuovo immobile</button>
        </div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {rentals.map((r) => {
            const econ = rentalEconomics(r, opDiRental(r, operations));
            const cfMensile = econ.cashFlowAnnuoNetto / 12;
            return (
              <div key={r.id} onClick={() => onOpen(r.id)} className="qg-panel" style={{ padding: 16, cursor: "pointer", display: "flex", alignItems: "center", gap: 16 }}>
                <div style={{ width: 40, height: 40, borderRadius: 9, background: "var(--accent-soft)", display: "flex", alignItems: "center", justifyContent: "center", border: "1px solid var(--accent)", flexShrink: 0 }}>
                  <Home size={18} color="var(--accent)" />
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ fontSize: 14, fontWeight: 600 }}>{r.nome || "Immobile senza nome"}</span>
                    <span className="qg-badge" style={{ fontSize: 10, padding: "2px 7px", borderRadius: 5, background: "var(--panel-2)", color: "var(--text-muted)" }}>{r.tipoAffitto === "lungo" ? "Affitto lungo" : "Affitto breve"}</span>
                    {!r.attivo && <span className="qg-badge" style={{ fontSize: 10, padding: "2px 7px", borderRadius: 5, background: "var(--panel-2)", color: "var(--text-dim)" }}>Non attivo</span>}
                  </div>
                  <div style={{ fontSize: 12, color: "var(--text-dim)", marginTop: 2, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{r.indirizzo || "—"}</div>
                </div>
                <div style={{ textAlign: "right", flexShrink: 0 }}>
                  <div className="qg-label" style={{ marginBottom: 2 }}>Rend. netto</div>
                  <div className="qg-mono" style={{ fontSize: 14, fontWeight: 600, color: "var(--accent)" }}>{fmtPct(econ.rendimentoNetto)}</div>
                </div>
                <div style={{ textAlign: "right", flexShrink: 0, minWidth: 110 }}>
                  <div className="qg-label" style={{ marginBottom: 2 }}>Cash flow / mese</div>
                  <div className="qg-mono" style={{ fontSize: 14, fontWeight: 600, color: cfMensile >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(cfMensile)}</div>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function RentalDetail({ rental, liquidity, operations, rentals, onEdit, onBack }) {
  const econ = rentalEconomics(rental, opDiRental(rental, operations));
  const cfMensile = econ.cashFlowAnnuoNetto / 12;
  const ric = rental.ricavi || {};

  // Impatto sulla liquidità: minimo proiettato considerando anche questo immobile.
  const soglia = Number(liquidity.soglioMinimo) || 0;
  const projection = useMemo(() => computeGlobalProjection(liquidity, operations, rentals, 60), [liquidity, operations, rentals]);
  const minRow = projection.reduce((min, r) => (r.balance < min.balance ? r : min), projection[0] || { balance: 0, month: currentYm() });
  const sostenibile = minRow.balance >= soglia;

  return (
    <div style={{ padding: 24, display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12 }}>
        <div style={{ display: "flex", alignItems: "flex-start", gap: 12 }}>
          <button className="qg-btn qg-btn-ghost" onClick={onBack} style={{ padding: 8 }}><ChevronLeft size={16} /></button>
          <div>
            <div className="qg-display" style={{ fontSize: 20, fontWeight: 600, display: "flex", alignItems: "center", gap: 10 }}>
              {rental.nome || "Immobile senza nome"}
              <span className="qg-badge" style={{ fontSize: 11, padding: "3px 9px", borderRadius: 6, background: "var(--panel-2)", color: "var(--text-muted)" }}>{rental.tipoAffitto === "lungo" ? "Affitto lungo" : "Affitto breve"}</span>
            </div>
            <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginTop: 3 }}>{rental.indirizzo || "—"}</div>
          </div>
        </div>
        <button className="qg-btn qg-btn-ghost" onClick={onEdit}><Pencil size={14} /> Modifica</button>
      </div>

      <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
        <RentalMetric label="Ricavo annuo lordo" value={fmtCompact(econ.ricavoAnnuoLordo)} tone="accent" sub={rental.tipoAffitto === "lungo" ? `canone ${fmtEUR(Number(ric.canoneMensile) || 0)}/mese` : `${fmtEUR(Number(ric.tariffaNotte) || 0)}/notte · ${Number(ric.occupazionePct) || 0}% occ.`} />
        <RentalMetric label="Cash flow netto / mese" value={fmtEURSigned(cfMensile)} tone={cfMensile >= 0 ? "positive" : "negative"} sub={`${fmtEURSigned(econ.cashFlowAnnuoNetto)} / anno`} />
        <RentalMetric label="Rendimento lordo" value={fmtPct(econ.rendimentoLordo)} sub="ricavo lordo / investimento" />
        <RentalMetric label="Rendimento netto" value={fmtPct(econ.rendimentoNetto)} tone="accent" sub="al netto dei costi di gestione" />
        <RentalMetric label="Cash-on-cash" value={fmtPct(econ.cashOnCash)} sub="su capitale proprio investito" />
      </div>

      <div className="qg-panel" style={{ padding: 20 }}>
        <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 14 }}>Conto economico annuo</div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <tbody>
            <tr style={{ borderBottom: "1px solid var(--border-soft)" }}>
              <td style={{ padding: "8px 4px" }}>Ricavo affitti (lordo)</td>
              <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", color: "var(--positive)" }}>{fmtEURSigned(econ.ricavoAnnuoLordo)}</td>
            </tr>
            <tr style={{ borderBottom: "1px solid var(--border-soft)" }}>
              <td style={{ padding: "8px 4px" }}>Costi di gestione ricorrenti</td>
              <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", color: "var(--negative)" }}>{fmtEURSigned(-econ.costiRicorrentiAnnui)}</td>
            </tr>
            <tr style={{ borderBottom: "1px solid var(--border-soft)" }}>
              <td style={{ padding: "8px 4px", fontWeight: 600 }}>Reddito operativo netto (NOI)</td>
              <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", fontWeight: 600 }}>{fmtEURSigned(econ.noiAnnuo)}</td>
            </tr>
            {econ.rataMutuoAnnua > 0 && (
              <tr style={{ borderBottom: "1px solid var(--border-soft)" }}>
                <td style={{ padding: "8px 4px" }}>Rata mutuo a regime (capitale + interessi)</td>
                <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", color: "var(--negative)" }}>{fmtEURSigned(-econ.rataMutuoAnnua)}</td>
              </tr>
            )}
            <tr>
              <td style={{ padding: "10px 4px", fontWeight: 700 }}>Cash flow netto annuo{econ.mesiPre > 0 ? " (a regime)" : ""}</td>
              <td className="qg-mono" style={{ padding: "10px 4px", textAlign: "right", fontWeight: 700, color: econ.cashFlowAnnuoNetto >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(econ.cashFlowAnnuoNetto)}</td>
            </tr>
            {econ.mesiPre > 0 && (
              <tr>
                <td style={{ padding: "8px 4px", color: "var(--text-muted)", fontSize: 12.5 }}>Cash flow durante il pre-ammortamento (primi {econ.mesiPre} mesi, soli interessi)</td>
                <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", color: econ.cashFlowAnnuoPre >= 0 ? "var(--positive)" : "var(--negative)", fontSize: 12.5 }}>{fmtEURSigned(econ.cashFlowAnnuoPre)}</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      <div className="qg-panel" style={{ padding: 16, background: sostenibile ? "var(--positive-soft)" : "var(--warning-soft)", border: `1px solid ${sostenibile ? "var(--positive)" : "var(--warning)"}`, display: "flex", gap: 10, alignItems: "flex-start" }}>
        {sostenibile ? <CheckCircle2 size={16} color="var(--positive)" style={{ marginTop: 2, flexShrink: 0 }} /> : <AlertTriangle size={16} color="var(--warning)" style={{ marginTop: 2, flexShrink: 0 }} />}
        <div style={{ fontSize: 13 }}>
          {sostenibile ? (
            <>Con questo immobile incluso, la liquidità QG resta sopra la soglia di sicurezza: minimo proiettato di <b>{fmtEUR(minRow.balance)}</b> a {monthLabelLong(minRow.month)}.</>
          ) : (
            <>Attenzione: includendo questo immobile, la liquidità QG proiettata scende a <b>{fmtEUR(minRow.balance)}</b> a {monthLabelLong(minRow.month)}, sotto la soglia di sicurezza di {fmtEUR(soglia)}.</>
          )}
        </div>
      </div>

      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <div className="qg-panel" style={{ padding: 20, flex: 1, minWidth: 280 }}>
          <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 12 }}>Investimento iniziale</div>
          {(rental.costiUnaTantum || []).length === 0 ? (
            <div style={{ fontSize: 12.5, color: "var(--text-dim)" }}>Nessun costo una tantum inserito.</div>
          ) : (
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
              <tbody>
                {rental.costiUnaTantum.map((c) => (
                  <tr key={c.id} style={{ borderBottom: "1px solid var(--border-soft)" }}>
                    <td style={{ padding: "6px 4px" }}>{c.categoria}</td>
                    <td style={{ padding: "6px 4px", color: "var(--text-dim)", fontSize: 11 }}>{dateLabel(c.data)}</td>
                    <td className="qg-mono" style={{ padding: "6px 4px", textAlign: "right" }}>{fmtEUR(Number(c.importo) || 0)}</td>
                  </tr>
                ))}
                <tr>
                  <td style={{ padding: "8px 4px", fontWeight: 700 }} colSpan={2}>Totale</td>
                  <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", fontWeight: 700 }}>{fmtEUR(econ.investimentoTotale)}</td>
                </tr>
              </tbody>
            </table>
          )}
          {econ.capitaleFinanziato > 0 && (
            <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginTop: 10 }}>
              Di cui {fmtEUR(econ.capitaleFinanziato)} finanziati con mutuo · {fmtEUR(econ.capitaleProprio)} capitale proprio.
            </div>
          )}
        </div>

        <div className="qg-panel" style={{ padding: 20, flex: 1, minWidth: 280 }}>
          <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 12 }}>Costi di gestione ricorrenti</div>
          {(rental.costiRicorrenti || []).length === 0 ? (
            <div style={{ fontSize: 12.5, color: "var(--text-dim)" }}>Nessun costo ricorrente inserito.</div>
          ) : (
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
              <tbody>
                {rental.costiRicorrenti.map((c) => (
                  <tr key={c.id} style={{ borderBottom: "1px solid var(--border-soft)" }}>
                    <td style={{ padding: "6px 4px" }}>{c.categoria}</td>
                    <td className="qg-mono" style={{ padding: "6px 4px", textAlign: "right" }}>{fmtEUR(Number(c.importoMensile) || 0)}/mese</td>
                  </tr>
                ))}
                <tr>
                  <td style={{ padding: "8px 4px", fontWeight: 700 }}>Totale annuo</td>
                  <td className="qg-mono" style={{ padding: "8px 4px", textAlign: "right", fontWeight: 700 }}>{fmtEUR(econ.costiRicorrentiAnnui)}</td>
                </tr>
              </tbody>
            </table>
          )}
        </div>
      </div>

      {rental.note && (
        <div className="qg-panel" style={{ padding: 16 }}>
          <div className="qg-label" style={{ marginBottom: 6 }}>Note</div>
          <div style={{ fontSize: 13, color: "var(--text-muted)", whiteSpace: "pre-wrap" }}>{rental.note}</div>
        </div>
      )}
    </div>
  );
}

function RentalEditor({ rental, liquidity, operations = [], onSave, onCancel, onDelete }) {
  const [r, setR] = useState(rental);
  const update = (patch) => setR((prev) => ({ ...prev, ...patch }));
  const updateRicavi = (patch) => setR((prev) => ({ ...prev, ricavi: { ...prev.ricavi, ...patch } }));
  const updateMutuo = (patch) => setR((prev) => ({ ...prev, mutuo: { ...prev.mutuo, ...patch } }));

  const addUnaTantum = () => update({ costiUnaTantum: [...(r.costiUnaTantum || []), { id: uid(), categoria: CAT_RENTAL_ONEOFF[0], importo: "", data: todayStr() }] });
  const updUnaTantum = (id, patch) => update({ costiUnaTantum: r.costiUnaTantum.map((c) => (c.id === id ? { ...c, ...patch } : c)) });
  const delUnaTantum = (id) => update({ costiUnaTantum: r.costiUnaTantum.filter((c) => c.id !== id) });

  const addRicorrente = () => update({ costiRicorrenti: [...(r.costiRicorrenti || []), { id: uid(), categoria: CAT_RENTAL_RICORRENTE[0], importoMensile: "", meseInizio: r.meseInizioLocazione, meseFine: "" }] });
  const updRicorrente = (id, patch) => update({ costiRicorrenti: r.costiRicorrenti.map((c) => (c.id === id ? { ...c, ...patch } : c)) });
  const delRicorrente = (id) => update({ costiRicorrenti: r.costiRicorrenti.filter((c) => c.id !== id) });

  const opCollegata = opDiRental(r, operations);
  const econ = rentalEconomics(r, opCollegata);
  // Operazioni marcate "a reddito" che questo immobile può agganciare per ereditarne costi e mutuo.
  const operazioniAReddito = (operations || []).filter((o) => o.destinazione === "reddito");

  return (
    <div style={{ padding: 24, display: "flex", flexDirection: "column", gap: 20, maxWidth: 920 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12 }}>
        <div className="qg-display" style={{ fontSize: 19, fontWeight: 600 }}>{rental.nome ? `Modifica: ${rental.nome}` : "Nuovo immobile a reddito"}</div>
        <div style={{ display: "flex", gap: 8 }}>
          {onDelete && <button className="qg-btn qg-btn-ghost qg-btn-danger" onClick={() => onDelete(r.id)}><Trash2 size={14} /> Elimina</button>}
          <button className="qg-btn qg-btn-ghost" onClick={onCancel}>Annulla</button>
          <button className="qg-btn qg-btn-primary" onClick={() => onSave(r)}><Save size={14} /> Salva</button>
        </div>
      </div>

      <div className="qg-panel" style={{ padding: 20 }}>
        <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 14 }}>Anagrafica</div>
        <div style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr 1fr", gap: 14 }}>
          <Field label="Nome / riferimento"><input className="qg-input" style={{ fontFamily: "Inter" }} value={r.nome} onChange={(e) => update({ nome: e.target.value })} placeholder="Es. Bilocale Duomo" /></Field>
          <Field label="Tipo di affitto">
            <select className="qg-input" value={r.tipoAffitto} onChange={(e) => update({ tipoAffitto: e.target.value })}>
              <option value="breve">Affitto breve (turistico)</option>
              <option value="lungo">Affitto lungo (canone mensile)</option>
            </select>
          </Field>
          <Field label="Inizio locazione"><input type="month" className="qg-input" value={r.meseInizioLocazione} onChange={(e) => update({ meseInizioLocazione: e.target.value })} /></Field>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 14, marginTop: 14 }}>
          <Field label="Indirizzo"><input className="qg-input" style={{ fontFamily: "Inter" }} value={r.indirizzo} onChange={(e) => update({ indirizzo: e.target.value })} placeholder="Città, zona, tipologia" /></Field>
          <Field label="In locazione">
            <select className="qg-input" value={r.attivo ? "si" : "no"} onChange={(e) => update({ attivo: e.target.value === "si" })}>
              <option value="si">Sì, genera ricavi</option>
              <option value="no">No (non ancora / sfitto)</option>
            </select>
          </Field>
        </div>
        <div style={{ marginTop: 14 }}>
          <Field label="Acquistato tramite operazione">
            <select className="qg-input" value={r.operazioneId || ""} onChange={(e) => update({ operazioneId: e.target.value })}>
              <option value="">Nessun collegamento — costi e mutuo inseriti qui</option>
              {operazioniAReddito.map((o) => <option key={o.id} value={o.id}>{o.nome || "Operazione senza nome"}</option>)}
            </select>
          </Field>
          <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginTop: 6 }}>
            {opCollegata
              ? <>Collegato a <b>{opCollegata.nome || "operazione"}</b>: costo d'acquisto ({fmtEUR(econ.investimentoOperazione)}) e mutuo sono presi da lì e non vanno reinseriti. Qui gestisci ricavi e costi di gestione.</>
              : operazioniAReddito.length === 0
                ? <>Per collegare un'operazione, impostane una con destinazione "A reddito" nella sezione Operazioni.</>
                : <>Collegando un'operazione, costo d'acquisto e mutuo vengono ereditati da quella scheda ed esclusi da qui, così non vengono contati due volte.</>}
          </div>
        </div>
      </div>

      <div className="qg-panel" style={{ padding: 20 }}>
        <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 14 }}>Ricavi stimati</div>
        {r.tipoAffitto === "lungo" ? (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr", gap: 14 }}>
            <Field label="Canone mensile (€)"><input type="number" className="qg-input" value={r.ricavi.canoneMensile} onChange={(e) => updateRicavi({ canoneMensile: e.target.value })} placeholder="1200" /></Field>
            <div style={{ alignSelf: "end", fontSize: 12.5, color: "var(--text-dim)", paddingBottom: 10 }}>Ricavo annuo stimato: <b className="qg-mono">{fmtEUR(econ.ricavoAnnuoLordo)}</b></div>
          </div>
        ) : (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 2fr", gap: 14 }}>
            <Field label="Tariffa media a notte (€)"><input type="number" className="qg-input" value={r.ricavi.tariffaNotte} onChange={(e) => updateRicavi({ tariffaNotte: e.target.value })} placeholder="120" /></Field>
            <Field label="Occupazione media (%)"><input type="number" className="qg-input" value={r.ricavi.occupazionePct} onChange={(e) => updateRicavi({ occupazionePct: e.target.value })} placeholder="65" /></Field>
            <div style={{ alignSelf: "end", fontSize: 12.5, color: "var(--text-dim)", paddingBottom: 10 }}>Ricavo annuo stimato: <b className="qg-mono">{fmtEUR(econ.ricavoAnnuoLordo)}</b> <span style={{ fontSize: 11 }}>(tariffa × 365 × occupazione)</span></div>
          </div>
        )}
      </div>

      <div className="qg-panel" style={{ padding: 20, display: opCollegata ? "none" : "block" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
          <div className="qg-label" style={{ fontSize: 12, marginBottom: 0, color: "var(--accent)" }}>Investimento iniziale (costi una tantum)</div>
          <button className="qg-btn qg-btn-ghost" onClick={addUnaTantum}><Plus size={14} /> Aggiungi costo</button>
        </div>
        <div style={{ fontSize: 12, color: "var(--text-dim)", marginBottom: 14 }}>Acquisto, ristrutturazione, arredo, notaio: tutto ciò che serve per mettere a reddito l'immobile.</div>
        {(r.costiUnaTantum || []).length === 0 && <div style={{ fontSize: 13, color: "var(--text-dim)" }}>Nessun costo inserito.</div>}
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {(r.costiUnaTantum || []).map((c) => (
            <div key={c.id} style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr 1fr auto", gap: 10, alignItems: "end" }}>
              <Field label="Categoria">
                <select className="qg-input" value={c.categoria} onChange={(e) => updUnaTantum(c.id, { categoria: e.target.value })}>
                  {CAT_RENTAL_ONEOFF.map((cat) => <option key={cat} value={cat}>{cat}</option>)}
                </select>
              </Field>
              <Field label="Importo (€)"><input type="number" className="qg-input" value={c.importo} onChange={(e) => updUnaTantum(c.id, { importo: e.target.value })} /></Field>
              <Field label={(liquidity?.dataRiferimento && c.data && c.data <= liquidity.dataRiferimento) ? "Data · già nel saldo" : "Data"}>
                <input type="date" className="qg-input" value={c.data} onChange={(e) => updUnaTantum(c.id, { data: e.target.value })}
                  style={(liquidity?.dataRiferimento && c.data && c.data <= liquidity.dataRiferimento) ? { borderStyle: "dashed", opacity: 0.7 } : undefined}
                  title={(liquidity?.dataRiferimento && c.data && c.data <= liquidity.dataRiferimento) ? "Datata entro la riconciliazione dei saldi: si considera già pagata e non viene più proiettata. Se la spesa è slittata, sposta la data in avanti." : undefined} />
              </Field>
              <button className="qg-btn qg-btn-ghost qg-btn-danger" onClick={() => delUnaTantum(c.id)} style={{ padding: 8 }}><Trash2 size={14} /></button>
            </div>
          ))}
        </div>
      </div>

      <div className="qg-panel" style={{ padding: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
          <div className="qg-label" style={{ fontSize: 12, marginBottom: 0, color: "var(--accent)" }}>Costi di gestione ricorrenti (mensili)</div>
          <button className="qg-btn qg-btn-ghost" onClick={addRicorrente}><Plus size={14} /> Aggiungi costo</button>
        </div>
        <div style={{ fontSize: 12, color: "var(--text-dim)", marginBottom: 14 }}>Property management, pulizie, utenze, IMU, assicurazione, condominio, commissioni piattaforme. Lascia vuoto "A" se prosegue indefinitamente.</div>
        {(r.costiRicorrenti || []).length === 0 && <div style={{ fontSize: 13, color: "var(--text-dim)" }}>Nessun costo ricorrente inserito.</div>}
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {(r.costiRicorrenti || []).map((c) => (
            <div key={c.id} style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr 0.9fr 0.9fr 0.8fr auto", gap: 10, alignItems: "end" }}>
              <Field label="Categoria">
                <select className="qg-input" value={c.categoria} onChange={(e) => updRicorrente(c.id, { categoria: e.target.value })}>
                  {CAT_RENTAL_RICORRENTE.map((cat) => <option key={cat} value={cat}>{cat}</option>)}
                </select>
              </Field>
              <Field label="€ / mese"><input type="number" className="qg-input" value={c.importoMensile} onChange={(e) => updRicorrente(c.id, { importoMensile: e.target.value })} /></Field>
              <Field label="Da"><input type="month" className="qg-input" value={c.meseInizio} onChange={(e) => updRicorrente(c.id, { meseInizio: e.target.value })} /></Field>
              <Field label="A (opz.)"><input type="month" className="qg-input" value={c.meseFine} onChange={(e) => updRicorrente(c.id, { meseFine: e.target.value })} /></Field>
              <Field label="Giorno add."><input type="number" min="1" max="31" className="qg-input" value={c.giornoAddebito ?? ""} onChange={(e) => updRicorrente(c.id, { giornoAddebito: e.target.value })} placeholder="fine mese" title="Giorno del mese in cui il costo viene addebitato. Serve a non riconteggiarlo se è già uscito prima della data di aggiornamento dei saldi." /></Field>
              <button className="qg-btn qg-btn-ghost qg-btn-danger" onClick={() => delRicorrente(c.id)} style={{ padding: 8 }}><Trash2 size={14} /></button>
            </div>
          ))}
        </div>
      </div>

      <div className="qg-panel" style={{ padding: 20, display: opCollegata ? "none" : "block" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
          <input type="checkbox" checked={!!r.mutuo.attivo} onChange={(e) => updateMutuo({ attivo: e.target.checked })} id="mutuo-attivo" />
          <label htmlFor="mutuo-attivo" className="qg-label" style={{ fontSize: 12, marginBottom: 0, color: "var(--accent)", cursor: "pointer" }}>Mutuo / finanziamento sull'immobile</label>
        </div>
        {r.mutuo.attivo && (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr", gap: 14 }}>
              <Field label="Importo (€)"><input type="number" className="qg-input" value={r.mutuo.importo} onChange={(e) => updateMutuo({ importo: e.target.value })} /></Field>
              <Field label="Tasso annuo (%)"><input type="number" step="0.1" className="qg-input" value={r.mutuo.tassoAnnuo} onChange={(e) => updateMutuo({ tassoAnnuo: e.target.value })} /></Field>
              <Field label="Durata (mesi)"><input type="number" className="qg-input" value={r.mutuo.durataMesi} onChange={(e) => updateMutuo({ durataMesi: e.target.value })} placeholder="240" /></Field>
              <Field label="Pre-ammort. (mesi)"><input type="number" min="0" className="qg-input" value={r.mutuo.mesiPreammortamento ?? ""} onChange={(e) => updateMutuo({ mesiPreammortamento: e.target.value })} placeholder="0" title="Mesi iniziali in cui paghi solo la quota interessi, senza restituire capitale. Sono compresi nella durata totale." /></Field>
              <Field label="Erogazione"><input type="month" className="qg-input" value={r.mutuo.meseErogazione} onChange={(e) => updateMutuo({ meseErogazione: e.target.value })} /></Field>
            </div>
            <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginTop: 10 }}>
              {econ.mesiPre > 0 ? (
                <>Primi <b>{econ.mesiPre} mesi</b> (pre-ammortamento, soli interessi): <b className="qg-mono">{fmtEUR(econ.rataPre)}</b>/mese · poi rata piena su {econ.mesiAmmortamento} mesi: <b className="qg-mono">{fmtEUR(econ.rataPiena)}</b>/mese</>
              ) : (
                <>Rata mensile stimata (ammortamento francese): <b className="qg-mono">{fmtEUR(econ.rataPiena)}</b>/mese</>
              )}
            </div>
          </>
        )}
      </div>

      <div className="qg-panel" style={{ padding: 20 }}>
        <Field label="Note"><textarea className="qg-input" style={{ fontFamily: "Inter", minHeight: 70, resize: "vertical" }} value={r.note} onChange={(e) => update({ note: e.target.value })} placeholder="Annotazioni libere sull'immobile" /></Field>
      </div>
    </div>
  );
}

/* ============================================================
   BILANCIO GESTIONALE (due rami: Operazioni / Smart Rent)
   ============================================================ */
function BilancioTooltip({ active, payload, label }) {
  if (!active || !payload || !payload.length) return null;
  const row = payload[0].payload;
  return (
    <div className="qg-panel qg-mono" style={{ padding: "10px 12px", fontSize: 12, border: "1px solid var(--border)" }}>
      <div style={{ color: "var(--text-muted)", marginBottom: 4 }}>{row.mese}</div>
      <div style={{ color: "var(--accent)" }}>Operazioni: {fmtEUR(row.opE)}</div>
      <div style={{ color: "var(--positive)" }}>Smart Rent: {fmtEUR(row.srE)}</div>
      <div style={{ color: "var(--text)", fontWeight: 600, marginTop: 3 }}>Fatturato: {fmtEUR(row.fatturato)}</div>
      <div style={{ color: row.utile >= 0 ? "var(--positive)" : "var(--negative)" }}>Utile: {fmtEURSigned(row.utile)}</div>
    </div>
  );
}

function BilancioUtileTooltip({ active, payload }) {
  if (!active || !payload || !payload.length) return null;
  const row = payload[0].payload;
  return (
    <div className="qg-panel qg-mono" style={{ padding: "10px 12px", fontSize: 12, border: "1px solid var(--border)" }}>
      <div style={{ color: "var(--text-muted)", marginBottom: 4 }}>{row.mese}</div>
      <div style={{ color: "var(--accent)" }}>Operazioni: {fmtEURSigned(row.opUtile)}</div>
      <div style={{ color: "var(--positive)" }}>Smart Rent: {fmtEURSigned(row.srUtile)}</div>
      <div style={{ color: row.utile >= 0 ? "var(--positive)" : "var(--negative)", fontWeight: 600, marginTop: 3 }}>Totale: {fmtEURSigned(row.utile)}</div>
    </div>
  );
}

function BilancioView({ bilancio, setBilancio }) {
  const [anno, setAnno] = useState(new Date().getFullYear());
  const annoStr = String(anno);
  const data = useMemo(() => bilancioAnnuale(bilancio, annoStr), [bilancio, annoStr]);

  const updCella = (meseIdx, ramo, campo, valore) => {
    const k = mkKey(meseIdx);
    setBilancio((prev) => {
      const anni = { ...(prev.anni || {}) };
      const mesi = { ...(anni[annoStr] || {}) };
      // Assicura che tutti i mesi esistano quando si tocca l'anno per la prima volta.
      for (let i = 0; i < 12; i++) { const kk = mkKey(i); if (!mesi[kk]) mesi[kk] = emptyMeseBilancio(); }
      const cella = { ...mesi[k], [ramo]: { ...mesi[k][ramo], [campo]: valore } };
      mesi[k] = cella;
      anni[annoStr] = mesi;
      return { ...prev, anni };
    });
  };

  const inputCella = (meseIdx, ramo, campo) => {
    const cell = (bilancio?.anni?.[annoStr]?.[mkKey(meseIdx)]) || emptyMeseBilancio();
    return (
      <input
        type="number"
        className="qg-input qg-mono"
        style={{ padding: "5px 6px", textAlign: "right", fontSize: 12 }}
        value={cell[ramo][campo]}
        onChange={(e) => updCella(meseIdx, ramo, campo, e.target.value)}
        placeholder="0"
      />
    );
  };

  return (
    <div style={{ padding: 24, display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
        <div>
          <div className="qg-display" style={{ fontSize: 20, fontWeight: 600 }}>Bilancio gestionale</div>
          <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginTop: 2 }}>Entrate e uscite dei due rami, inserite a mano mese per mese. Fatturato e utile calcolati automaticamente.</div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <button className="qg-btn qg-btn-ghost" onClick={() => setAnno((a) => a - 1)} style={{ padding: 8 }}><ChevronLeft size={16} /></button>
          <span className="qg-mono" style={{ fontSize: 16, fontWeight: 600, minWidth: 56, textAlign: "center" }}>{anno}</span>
          <button className="qg-btn qg-btn-ghost" onClick={() => setAnno((a) => a + 1)} style={{ padding: 8 }}><ChevronRight size={16} /></button>
        </div>
      </div>

      <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
        <StatCard icon={CircleDollarSign} label={`Fatturato ${anno}`} value={fmtCompact(data.fatturato)} sub="entrate totali dei due rami" />
        <StatCard icon={TrendingDown} label={`Uscite ${anno}`} value={fmtCompact(data.uscite)} sub="costi totali dei due rami" />
        <StatCard icon={PiggyBank} label={`Utile ${anno}`} value={fmtCompact(data.utile)} tone={data.utile >= 0 ? "positive" : "negative"} sub={data.fatturato > 0 ? `margine ${fmtPct(data.margine)}` : undefined} />
      </div>

      <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
        <div className="qg-panel" style={{ padding: 16, flex: 1, minWidth: 240 }}>
          <div className="qg-label" style={{ marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 9, height: 9, borderRadius: 2, background: "var(--accent)" }} /> Lato Operazioni</div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 3 }}><span style={{ color: "var(--text-dim)" }}>Fatturato</span><span className="qg-mono">{fmtEUR(data.tot.operazioni.entrate)}</span></div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 3 }}><span style={{ color: "var(--text-dim)" }}>− Uscite</span><span className="qg-mono">{fmtEUR(data.tot.operazioni.uscite)}</span></div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 3 }}><span style={{ color: "var(--text-dim)" }}>− Capitale restituito</span><span className="qg-mono">{fmtEUR(data.tot.operazioni.restituzione)}</span></div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, fontWeight: 600, paddingTop: 5, borderTop: "1px solid var(--border-soft)" }}><span>Utile</span><span className="qg-mono" style={{ color: data.opUtile >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(data.opUtile)}</span></div>
        </div>
        <div className="qg-panel" style={{ padding: 16, flex: 1, minWidth: 240 }}>
          <div className="qg-label" style={{ marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 9, height: 9, borderRadius: 2, background: "var(--positive)" }} /> Smart Rent Solution</div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 3 }}><span style={{ color: "var(--text-dim)" }}>Fatturato</span><span className="qg-mono">{fmtEUR(data.tot.smartRent.entrate)}</span></div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 3 }}><span style={{ color: "var(--text-dim)" }}>Uscite</span><span className="qg-mono">{fmtEUR(data.tot.smartRent.uscite)}</span></div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, fontWeight: 600, paddingTop: 5, borderTop: "1px solid var(--border-soft)" }}><span>Utile</span><span className="qg-mono" style={{ color: data.srUtile >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(data.srUtile)}</span></div>
        </div>
      </div>

      <div className="qg-panel" style={{ padding: 20 }}>
        <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 14 }}>Fatturato mensile per ramo</div>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={data.perMese} margin={{ top: 6, right: 10, left: 0, bottom: 0 }}>
            <CartesianGrid stroke="var(--border-soft)" vertical={false} />
            <XAxis dataKey="mese" tick={{ fill: "var(--text-dim)", fontSize: 11 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
            <YAxis tick={{ fill: "var(--text-dim)", fontSize: 11 }} axisLine={false} tickLine={false} tickFormatter={fmtCompact} width={56} />
            <Tooltip content={<BilancioTooltip />} cursor={{ fill: "var(--panel-2)" }} />
            <Bar dataKey="opE" stackId="fat" fill="var(--accent)" radius={[0, 0, 0, 0]} />
            <Bar dataKey="srE" stackId="fat" fill="var(--positive)" radius={[3, 3, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
        <div style={{ display: "flex", gap: 18, marginTop: 8, fontSize: 11, color: "var(--text-dim)" }}>
          <span style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 10, height: 10, borderRadius: 2, background: "var(--accent)" }} /> Lato Operazioni</span>
          <span style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 10, height: 10, borderRadius: 2, background: "var(--positive)" }} /> Smart Rent Solution</span>
        </div>
      </div>

      <div className="qg-panel" style={{ padding: 20 }}>
        <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 14 }}>Utile mensile per ramo</div>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={data.perMese} margin={{ top: 6, right: 10, left: 0, bottom: 0 }}>
            <CartesianGrid stroke="var(--border-soft)" vertical={false} />
            <XAxis dataKey="mese" tick={{ fill: "var(--text-dim)", fontSize: 11 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
            <YAxis tick={{ fill: "var(--text-dim)", fontSize: 11 }} axisLine={false} tickLine={false} tickFormatter={fmtCompact} width={56} />
            <Tooltip content={<BilancioUtileTooltip />} cursor={{ fill: "var(--panel-2)" }} />
            <ReferenceLine y={0} stroke="var(--border)" />
            <Bar dataKey="opUtile" fill="var(--accent)" radius={[3, 3, 0, 0]} />
            <Bar dataKey="srUtile" fill="var(--positive)" radius={[3, 3, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
        <div style={{ display: "flex", gap: 18, marginTop: 8, fontSize: 11, color: "var(--text-dim)" }}>
          <span style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 10, height: 10, borderRadius: 2, background: "var(--accent)" }} /> Lato Operazioni</span>
          <span style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 10, height: 10, borderRadius: 2, background: "var(--positive)" }} /> Smart Rent Solution</span>
        </div>
      </div>

      <div className="qg-panel" style={{ padding: 20, overflowX: "auto" }}>
        <div className="qg-label" style={{ fontSize: 12, color: "var(--accent)", marginBottom: 6 }}>Dettaglio mensile {anno}</div>
        <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginBottom: 14 }}>La restituzione capitale è la quota di entrate operazioni che è denaro di terzi da rimborsare: non tocca il fatturato, ma viene sottratta all'utile. Utile = fatturato − uscite − restituzione.</div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, minWidth: 820 }}>
          <thead>
            <tr style={{ color: "var(--text-dim)", fontSize: 10.5, textTransform: "uppercase", letterSpacing: "0.03em" }}>
              <th style={{ textAlign: "left", padding: "0 8px 8px", fontWeight: 500 }}>Mese</th>
              <th style={{ textAlign: "right", padding: "0 8px 8px", fontWeight: 500, color: "var(--accent)" }}>Op. Entrate</th>
              <th style={{ textAlign: "right", padding: "0 8px 8px", fontWeight: 500, color: "var(--accent)" }}>Op. Restituz.</th>
              <th style={{ textAlign: "right", padding: "0 8px 8px", fontWeight: 500, color: "var(--accent)" }}>Op. Uscite</th>
              <th style={{ textAlign: "right", padding: "0 8px 8px", fontWeight: 500, color: "var(--positive)" }}>SR Entrate</th>
              <th style={{ textAlign: "right", padding: "0 8px 8px", fontWeight: 500, color: "var(--positive)" }}>SR Uscite</th>
              <th style={{ textAlign: "right", padding: "0 8px 8px", fontWeight: 500 }}>Fatturato</th>
              <th style={{ textAlign: "right", padding: "0 8px 8px", fontWeight: 500 }}>Utile</th>
            </tr>
          </thead>
          <tbody>
            {data.perMese.map((m) => (
              <tr key={m.idx} style={{ borderTop: "1px solid var(--border-soft)" }}>
                <td style={{ padding: "5px 8px", fontWeight: 500 }}>{m.mese}</td>
                <td style={{ padding: "3px 8px", width: 100 }}>{inputCella(m.idx, "operazioni", "entrate")}</td>
                <td style={{ padding: "3px 8px", width: 100 }}>{inputCella(m.idx, "operazioni", "restituzione")}</td>
                <td style={{ padding: "3px 8px", width: 100 }}>{inputCella(m.idx, "operazioni", "uscite")}</td>
                <td style={{ padding: "3px 8px", width: 100 }}>{inputCella(m.idx, "smartRent", "entrate")}</td>
                <td style={{ padding: "3px 8px", width: 100 }}>{inputCella(m.idx, "smartRent", "uscite")}</td>
                <td className="qg-mono" style={{ padding: "5px 8px", textAlign: "right" }}>{m.fatturato ? fmtEUR(m.fatturato) : "—"}</td>
                <td className="qg-mono" style={{ padding: "5px 8px", textAlign: "right", color: m.utile > 0 ? "var(--positive)" : m.utile < 0 ? "var(--negative)" : "var(--text-dim)" }}>{m.fatturato || m.uscite ? fmtEURSigned(m.utile) : "—"}</td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr style={{ borderTop: "2px solid var(--border)", fontWeight: 700 }}>
              <td style={{ padding: "8px" }}>Totale</td>
              <td className="qg-mono" style={{ padding: "8px", textAlign: "right" }}>{fmtEUR(data.tot.operazioni.entrate)}</td>
              <td className="qg-mono" style={{ padding: "8px", textAlign: "right" }}>{fmtEUR(data.tot.operazioni.restituzione)}</td>
              <td className="qg-mono" style={{ padding: "8px", textAlign: "right" }}>{fmtEUR(data.tot.operazioni.uscite)}</td>
              <td className="qg-mono" style={{ padding: "8px", textAlign: "right" }}>{fmtEUR(data.tot.smartRent.entrate)}</td>
              <td className="qg-mono" style={{ padding: "8px", textAlign: "right" }}>{fmtEUR(data.tot.smartRent.uscite)}</td>
              <td className="qg-mono" style={{ padding: "8px", textAlign: "right" }}>{fmtEUR(data.fatturato)}</td>
              <td className="qg-mono" style={{ padding: "8px", textAlign: "right", color: data.utile >= 0 ? "var(--positive)" : "var(--negative)" }}>{fmtEURSigned(data.utile)}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}

/* ============================================================
   CHIEDI A CLAUDE (analisi conversazionale sui dati)
   ============================================================ */
// Builds a compact, human-readable snapshot of the whole workspace, computed with the very
// same engine functions the UI uses, so the assistant reasons on the numbers the user sees.
function buildDataSnapshot(liquidity, operations, rentals, bilancio) {
  const L = [];
  const push = (s) => L.push(s);
  const eur = (n) => fmtEUR(Number(n) || 0);

  push(`Data odierna: ${dateLabel(todayStr())}`);

  // --- Liquidità ---
  push(`\n## LIQUIDITÀ`);
  const conti = (liquidity.conti && liquidity.conti.length) ? liquidity.conti : [];
  conti.forEach((c) => push(`- Conto "${c.nome || "senza nome"}": ${eur(c.saldo)}${Number(c.vincolato) > 0 ? ` (di cui ${eur(c.vincolato)} di terzi, non spendibili)` : ""}`));
  const totLiq = totaleDisponibile(liquidity);
  const totVinc = totaleVincolato(liquidity);
  if (totVinc > 0) push(`- Totale in banca: ${eur(totaleLiquidita(liquidity))}, di cui ${eur(totVinc)} denaro di terzi (quota proprietari gestiti da Smart Rent).`);
  push(`- LIQUIDITÀ DISPONIBILE (spendibile, base della proiezione): ${eur(totLiq)}`);
  push(`- Saldi aggiornati (riconciliati) al: ${dateLabel(liquidity.dataRiferimento)}. Tutto ciò che è datato fino a questa data è già dentro i saldi e non viene proiettato di nuovo.`);
  push(`- Soglia minima di sicurezza: ${eur(liquidity.soglioMinimo)}`);
  const costiFissiMensili = (liquidity.costiFissi?.ricorrenti || []).reduce((s, c) => {
    const now = currentYm();
    const attivo = c.meseInizio && monthDiff(c.meseInizio, now) >= 0 && (!c.meseFine || monthDiff(now, c.meseFine) >= 0);
    return attivo ? s + (Number(c.importoMensile) || 0) : s;
  }, 0);
  push(`- Costi fissi aziendali ricorrenti attuali: ${eur(costiFissiMensili)}/mese`);
  (liquidity.costiFissi?.ricorrenti || []).forEach((c) => {
    push(`  · ${c.categoria}: ${eur(c.importoMensile)}/mese, da ${c.meseInizio || "?"}${c.meseFine ? ` a ${c.meseFine}` : " (senza data di fine)"}${c.giornoAddebito ? `, addebito il giorno ${c.giornoAddebito}` : ""}`);
  });

  // --- Proiezione ---
  const proj = computeGlobalProjection(liquidity, operations, rentals, 60);
  const soglia = Number(liquidity.soglioMinimo) || 0;
  const min = proj.reduce((m, r) => (r.balance < m.balance ? r : m), proj[0] || { balance: totLiq, month: currentYm() });
  const tra12 = proj[Math.min(11, proj.length - 1)];
  push(`\n## PROIEZIONE LIQUIDITÀ (60 mesi, include operazioni, immobili a reddito, costi fissi, debiti/crediti)`);
  push(`- Saldo fra 12 mesi: ${eur(tra12?.balance)}`);
  push(`- Minimo proiettato: ${eur(min.balance)} a ${monthLabelLong(min.month)}${min.balance < soglia ? " (SOTTO la soglia di sicurezza)" : ""}`);
  const primoNegativo = proj.find((r) => r.balance < soglia);
  if (primoNegativo) push(`- Primo mese sotto soglia: ${monthLabelLong(primoNegativo.month)}`);

  // --- Operazioni ---
  push(`\n## OPERAZIONI (compravendita)`);
  if (!operations.length) push(`- Nessuna operazione inserita.`);
  operations.forEach((op) => {
    const a = analyzeSustainability(op, liquidity, operations, rentals);
    const s = a.share;
    push(`- "${op.nome || "senza nome"}" [${op.stato}] ${op.indirizzo ? `(${op.indirizzo})` : ""}`);
    push(`  · Destinazione: ${op.destinazione === "reddito" ? "A REDDITO (tenuto e affittato, nessuna vendita)" : (op.frazionamento ? "flipping con frazionamento in più unità" : "flipping (rivendita)")}`);
    push(`  · Periodo: da ${op.meseInizio || "?"}${a.aReddito ? "" : ` a chiusura prevista ${a.meseChiusura || "non impostata"}`}`);
    push(`  · Fabbisogno iniziale: ${eur(s.fabbisognoIniziale)} | Copertura: liquidità QG ${eur(s.liquiditaQG)}, prestito ${eur(s.prestito)}, investitori ${eur(s.investitori)}${s.gap > 0 ? ` | SCOPERTO: ${eur(s.gap)}` : ""}`);
    if (!a.aReddito) {
      push(`  · Ricavo di vendita previsto: ${eur(a.ricavoTotale)}${a.vendite.length > 1 ? ` su ${a.vendite.length} unità: ${a.vendite.map((v) => `${v.descrizione} ${eur(v.importo)} a ${v.mese}`).join("; ")}` : ""}`);
    }
    push(`  · Margine lordo: ${eur(a.marginLordo)} | netto imposte: ${eur(a.marginNetto)} | quota QG: ${eur(a.marginQG)}`);
    if (a.investitoriDettaglio?.length) {
      a.investitoriDettaglio.forEach((i) => push(`  · Investitore ${i.nome}: capitale ${eur(i.importo)} (${fmtPct(i.quotaPct)}), utile netto ${eur(i.utileNetto)}`));
    }
    push(`  · Valutazione sostenibilità: ${a.verdict}${a.deficitLiquidita > 0 ? ` (deficit ${eur(a.deficitLiquidita)})` : ""}`);
  });

  // --- Immobili a reddito ---
  push(`\n## IMMOBILI A REDDITO (affitti)`);
  if (!rentals.length) push(`- Nessun immobile a reddito inserito.`);
  rentals.forEach((r) => {
    const e = rentalEconomics(r, opDiRental(r, operations));
    push(`- "${r.nome || "senza nome"}" [${r.tipoAffitto === "lungo" ? "affitto lungo" : "affitto breve"}${r.attivo ? "" : ", NON attivo"}] ${r.indirizzo ? `(${r.indirizzo})` : ""}`);
    push(`  · Ricavo annuo lordo: ${eur(e.ricavoAnnuoLordo)} | costi gestione: ${eur(e.costiRicorrentiAnnui)}/anno | rata mutuo: ${eur(e.rataMutuoAnnua)}/anno${e.mesiPre > 0 ? ` (primi ${e.mesiPre} mesi in pre-ammortamento a soli interessi: ${eur(e.rataPre)}/mese, poi ${eur(e.rataPiena)}/mese)` : ""}`);
    push(`  · Investimento totale: ${eur(e.investimentoTotale)} (capitale proprio ${eur(e.capitaleProprio)}, mutuo ${eur(e.capitaleFinanziato)})`);
    push(`  · Cash flow netto: ${eur(e.cashFlowAnnuoNetto)}/anno | Rend. lordo ${fmtPct(e.rendimentoLordo)} | Rend. netto ${fmtPct(e.rendimentoNetto)} | Cash-on-cash ${fmtPct(e.cashOnCash)}`);
  });

  // --- Debiti e crediti ---
  const partite = (liquidity.partite || []).filter((p) => !p.saldato);
  push(`\n## DEBITI E CREDITI APERTI`);
  if (!partite.length) push(`- Nessuna partita aperta.`);
  partite.forEach((p) => push(`- ${p.tipo === "credito" ? "CREDITO da incassare" : "DEBITO da rimborsare"}: ${eur(p.importo)} da/a "${p.controparte || "?"}" il ${dateLabel(p.data)}${p.descrizione ? ` — ${p.descrizione}` : ""}`));

  // --- Portafoglio titoli ---
  const port = portfolioSummary(liquidity.portafoglio, liquidity.fx?.usdToEur);
  push(`\n## PORTAFOGLIO TITOLI (riserva liquidabile, NON inclusa nella proiezione di cassa)`);
  if (!port.titoli.length) push(`- Nessun titolo.`);
  port.titoli.forEach((t) => push(`- ${t.nome || t.ticker || "titolo"}${t.ticker ? ` (${t.ticker})` : ""}: ${t.quantita} @ ${t.valuta}, valore ${eur(t.valoreAttuale)}, P/L ${eur(t.pl)} (${fmtPct(t.plPct)})`));
  if (port.titoli.length) push(`- Totale investito ${eur(port.investito)}, valore attuale ${eur(port.valoreAttuale)}, P/L ${eur(port.pl)} (${fmtPct(port.plPct)})`);

  // --- Scadenzario spese ---
  const spese = buildSpeseSchedule(operations, liquidity).filter((s) => s.data >= todayStr()).slice(0, 10);
  push(`\n## PROSSIME SPESE UNA TANTUM IN SCADENZA`);
  if (!spese.length) push(`- Nessuna spesa in programma.`);
  spese.forEach((s) => push(`- ${dateLabel(s.data)}: ${eur(s.importo)} — ${s.categoria} (${s.origine})`));

  // --- Bilancio ---
  const annoCorr = new Date().getFullYear();
  push(`\n## BILANCIO GESTIONALE (inserito manualmente)`);
  [annoCorr - 1, annoCorr].forEach((y) => {
    const b = bilancioAnnuale(bilancio, String(y));
    if (b.fatturato === 0 && b.uscite === 0) return;
    push(`- Anno ${y}: fatturato ${eur(b.fatturato)}, uscite ${eur(b.uscite)}, capitale restituito ${eur(b.tot.operazioni.restituzione)}, UTILE ${eur(b.utile)} (margine ${fmtPct(b.margine)})`);
    push(`  · Lato Operazioni: fatturato ${eur(b.tot.operazioni.entrate)}, uscite ${eur(b.tot.operazioni.uscite)}, utile ${eur(b.opUtile)}`);
    push(`  · Smart Rent Solution: fatturato ${eur(b.tot.smartRent.entrate)}, uscite ${eur(b.tot.smartRent.uscite)}, utile ${eur(b.srUtile)}`);
  });

  return L.join("\n");
}

const DOMANDE_SUGGERITE = [
  "Qual è il mese più critico per la cassa e perché?",
  "Quale operazione rende di più rispetto al capitale che immobilizza?",
  "Posso permettermi una nuova operazione da 300.000 €?",
  "Conviene di più il lato operazioni o gli affitti brevi?",
  "Cosa succede se una vendita slitta di 6 mesi?",
];

function AskClaudeView({ liquidity, operations, rentals, bilancio }) {
  const [messages, setMessages] = useState([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const scrollRef = useRef(null);

  useEffect(() => {
    if (scrollRef.current) scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
  }, [messages, loading]);

  const invia = async (testo) => {
    const domanda = (testo ?? input).trim();
    if (!domanda || loading) return;
    setInput("");
    setError("");
    const nuoviMessaggi = [...messages, { role: "user", content: domanda }];
    setMessages(nuoviMessaggi);
    setLoading(true);

    try {
      const snapshot = buildDataSnapshot(liquidity, operations, rentals, bilancio);
      const system = `Sei un analista finanziario esperto di real estate che affianca QG Real Estate, una società che fa operazioni di compravendita immobiliare e gestisce immobili a reddito (affitti brevi, ramo "Smart Rent Solution").

Rispondi SEMPRE in italiano, in modo diretto e concreto, come farebbe un consulente che conosce bene i numeri dell'azienda.

Ecco lo stato attuale dei dati inseriti nel gestionale:

${snapshot}

REGOLE IMPORTANTI:
- Basa le risposte SOLO su questi dati. Se un dato non c'è, dillo chiaramente invece di inventarlo.
- Quando fai un calcolo, mostra i passaggi essenziali così l'utente può verificarlo.
- Usa gli importi in euro con separatori leggibili.
- Sii sintetico: vai al punto, niente premesse lunghe. Usa elenchi solo se aiutano davvero.
- Se noti un rischio rilevante (cassa sotto soglia, operazione scoperta, margine negativo), segnalalo anche se non è stato chiesto.
- Ricorda: il portafoglio titoli è una riserva liquidabile ma NON è inclusa nella proiezione di cassa operativa. La "restituzione capitale" nel bilancio non riduce il fatturato ma abbatte l'utile.
- Non sei un consulente fiscale o legale abilitato: per decisioni fiscali o contrattuali importanti, ricorda di verificare con il commercialista.`;

      const data = await callClaude({
        max_tokens: 1000,
        system,
        messages: nuoviMessaggi.map((m) => ({ role: m.role, content: m.content })),
      });
      const testoRisposta = (data.content || [])
        .filter((b) => b.type === "text")
        .map((b) => b.text)
        .join("\n")
        .trim();

      setMessages([...nuoviMessaggi, { role: "assistant", content: testoRisposta || "Non ho ricevuto una risposta leggibile." }]);
    } catch (e) {
      setError("Non sono riuscito a contattare Claude. Riprova fra un momento.");
      setMessages(nuoviMessaggi);
    } finally {
      setLoading(false);
    }
  };

  const onKeyDown = (e) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); invia(); }
  };

  return (
    <div style={{ padding: 24, display: "flex", flexDirection: "column", gap: 16, height: "100%", maxWidth: 900 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
        <div>
          <div className="qg-display" style={{ fontSize: 20, fontWeight: 600, display: "flex", alignItems: "center", gap: 8 }}>
            <Sparkles size={18} color="var(--accent)" /> Chiedi a Claude
          </div>
          <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginTop: 2 }}>Fai domande sui tuoi dati: liquidità, operazioni, immobili a reddito, bilancio. Claude legge lo stato aggiornato del gestionale ad ogni domanda.</div>
        </div>
        {messages.length > 0 && (
          <button className="qg-btn qg-btn-ghost" onClick={() => { setMessages([]); setError(""); }} style={{ fontSize: 11.5 }}>Nuova conversazione</button>
        )}
      </div>

      <div ref={scrollRef} className="qg-panel qg-scroll" style={{ flex: 1, minHeight: 300, maxHeight: "58vh", overflowY: "auto", padding: 18, display: "flex", flexDirection: "column", gap: 14 }}>
        {messages.length === 0 && !loading && (
          <div style={{ margin: "auto 0", textAlign: "center", padding: "20px 10px" }}>
            <Sparkles size={24} color="var(--text-dim)" style={{ marginBottom: 10 }} />
            <div style={{ fontSize: 13.5, marginBottom: 4 }}>Chiedimi qualcosa sui tuoi numeri</div>
            <div style={{ fontSize: 12, color: "var(--text-dim)", marginBottom: 18 }}>Ho accesso a liquidità, operazioni, immobili, debiti/crediti, titoli e bilancio.</div>
            <div style={{ display: "flex", flexDirection: "column", gap: 7, maxWidth: 430, margin: "0 auto" }}>
              {DOMANDE_SUGGERITE.map((d) => (
                <button key={d} className="qg-btn qg-btn-ghost" style={{ fontSize: 12, textAlign: "left", justifyContent: "flex-start", padding: "9px 12px", lineHeight: 1.35 }} onClick={() => invia(d)}>{d}</button>
              ))}
            </div>
          </div>
        )}

        {messages.map((m, i) => (
          <div key={i} style={{ display: "flex", flexDirection: "column", alignItems: m.role === "user" ? "flex-end" : "flex-start" }}>
            <div className="qg-label" style={{ marginBottom: 4, fontSize: 10 }}>{m.role === "user" ? "Tu" : "Claude"}</div>
            <div
              className="qg-panel"
              style={{
                padding: "11px 14px", maxWidth: "88%", fontSize: 13.2, lineHeight: 1.55, whiteSpace: "pre-wrap",
                background: m.role === "user" ? "var(--accent-soft)" : "var(--panel-2)",
                border: `1px solid ${m.role === "user" ? "var(--accent)" : "var(--border-soft)"}`,
              }}
            >
              {m.content}
            </div>
          </div>
        ))}

        {loading && (
          <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-start" }}>
            <div className="qg-label" style={{ marginBottom: 4, fontSize: 10 }}>Claude</div>
            <div className="qg-panel" style={{ padding: "11px 14px", fontSize: 13, color: "var(--text-dim)", background: "var(--panel-2)", border: "1px solid var(--border-soft)" }}>Sto analizzando i dati…</div>
          </div>
        )}
      </div>

      {error && (
        <div className="qg-panel" style={{ padding: 11, background: "var(--negative-soft)", border: "1px solid var(--negative)", fontSize: 12.5, display: "flex", gap: 8, alignItems: "center" }}>
          <ShieldAlert size={14} color="var(--negative)" style={{ flexShrink: 0 }} /> {error}
        </div>
      )}

      <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
        <textarea
          className="qg-input"
          style={{ fontFamily: "Inter", flex: 1, minHeight: 46, maxHeight: 130, resize: "vertical", lineHeight: 1.4 }}
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={onKeyDown}
          placeholder="Es. Se compro un immobile da 400k a ottobre, la cassa regge?"
          disabled={loading}
        />
        <button className="qg-btn qg-btn-primary" onClick={() => invia()} disabled={loading || !input.trim()} style={{ height: 46, padding: "0 16px" }}>
          <SendHorizontal size={15} /> Invia
        </button>
      </div>
      <div style={{ fontSize: 11, color: "var(--text-dim)", marginTop: -6 }}>
        Claude ragiona sui dati che hai inserito e può commettere errori: verifica sempre i numeri prima di decidere. Le conversazioni non vengono salvate.
      </div>
    </div>
  );
}

/* ============================================================
   MAIN APP
   ============================================================ */
export default function QGRealEstateApp() {
  const [loading, setLoading] = useState(true);
  const [liquidity, setLiquidity] = useState(emptyLiquidity());
  const [operations, setOperations] = useState([]);
  const [rentals, setRentals] = useState([]);
  const [bilancio, setBilancio] = useState(emptyBilancio());
  const [tab, setTab] = useState("dashboard");
  const [filterStato, setFilterStato] = useState("tutte");
  const [openOpId, setOpenOpId] = useState(null);
  const [returnTab, setReturnTab] = useState("operazioni");
  const [editingOp, setEditingOp] = useState(null);
  const [openRentalId, setOpenRentalId] = useState(null);
  const [editingRental, setEditingRental] = useState(null);
  const [saveStatus, setSaveStatus] = useState("");
  const [importError, setImportError] = useState("");
  const loadedRef = useRef(false);
  const importInputRef = useRef(null);

  useEffect(() => {
    (async () => {
      try {
        const res = await storage.get("qg-real-estate-data");
        if (res && res.value) {
          const parsed = JSON.parse(res.value);
          if (parsed.liquidity) setLiquidity(normalizeLiquidity(parsed.liquidity));
          if (parsed.operations) setOperations(parsed.operations.map(normalizeOperation));
          if (parsed.rentals) setRentals(parsed.rentals.map(normalizeRental));
          if (parsed.bilancio) setBilancio(normalizeBilancio(parsed.bilancio));
        }
        loadedRef.current = true;
      } catch (e) {
        // Lettura fallita (rete o permessi): il salvataggio resta disabilitato, altrimenti
        // uno stato vuoto sovrascriverebbe i dati reali sul database.
        loadedRef.current = false;
        setSaveStatus("Errore di caricamento: dati non disponibili. Ricarica la pagina.");
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  useEffect(() => {
    if (loadedRef.current !== true) return;
    const t = setTimeout(async () => {
      try {
        await storage.set("qg-real-estate-data", JSON.stringify({ liquidity, operations, rentals, bilancio }));
        setSaveStatus("Salvato " + new Date().toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" }));
      } catch (e) {
        // Un altro socio ha salvato nel frattempo: non sovrascriviamo le sue modifiche.
        setSaveStatus(e instanceof ConflictError
          ? "Non salvato: dati modificati da un altro utente. Ricarica la pagina."
          : "Errore di salvataggio");
      }
    }, 600);
    return () => clearTimeout(t);
  }, [liquidity, operations, rentals, bilancio]);

  // Auto-refresh the USD→EUR rate once per day, after data has loaded (silent; manual button also exists).
  const fxCheckedRef = useRef(false);
  useEffect(() => {
    if (loading || fxCheckedRef.current) return;
    fxCheckedRef.current = true;
    const aggiornato = liquidity.fx?.aggiornato;
    if (aggiornato && aggiornato !== "manuale" && aggiornato === todayStr()) return; // già aggiornato oggi
    (async () => {
      try {
        const { usdToEur, date } = await fetchFxRate();
        setLiquidity((prev) => {
          // Non sovrascrivere un valore inserito a mano oggi.
          if (prev.fx?.aggiornato === "manuale") return prev;
          return { ...prev, fx: { usdToEur: String(usdToEur), aggiornato: date } };
        });
      } catch (e) { /* offline o bloccato: si resta sul valore esistente/manuale */ }
    })();
  }, [loading, liquidity.fx?.aggiornato]);

  const saveOperation = (op) => {
    setOperations((prev) => {
      const exists = prev.some((o) => o.id === op.id);
      return exists ? prev.map((o) => (o.id === op.id ? op : o)) : [...prev, op];
    });
    setEditingOp(null);
    setOpenOpId(op.id);
  };

  const deleteOperation = (id) => {
    setOperations((prev) => prev.filter((o) => o.id !== id));
    setEditingOp(null);
    setOpenOpId(null);
    setTab("operazioni");
  };

  const saveRental = (r) => {
    setRentals((prev) => {
      const exists = prev.some((x) => x.id === r.id);
      return exists ? prev.map((x) => (x.id === r.id ? r : x)) : [...prev, r];
    });
    setEditingRental(null);
    setOpenRentalId(r.id);
  };

  const deleteRental = (id) => {
    setRentals((prev) => prev.filter((x) => x.id !== id));
    setEditingRental(null);
    setOpenRentalId(null);
    setTab("immobili");
  };

  const seedDemo = () => {
    const demo = emptyOperation();
    demo.nome = "Via Torino 12";
    demo.indirizzo = "Milano, zona Navigli — trilocale da ristrutturare";
    demo.stato = "in_corso";
    demo.meseInizio = currentYm();
    demo.meseVenditaPrevista = addMonths(currentYm(), 9);
    demo.prezzoVenditaPrevisto = "480000";
    demo.costiUnaTantum = [
      { id: uid(), categoria: "Acquisto", importo: "300000", data: todayStr() },
      { id: uid(), categoria: "Ristrutturazione / Lavori", importo: "80000", data: toDateStr(addMonths(currentYm(), 2)) },
      { id: uid(), categoria: "Notaio e Imposte", importo: "20000", data: todayStr() },
    ];
    demo.costiRicorrenti = [{ id: uid(), categoria: "Gestione / Amministrazione", importoMensile: "500", meseInizio: currentYm(), meseFine: "" }];
    demo.copertura = {
      liquiditaQG: "150000",
      prestito: { attivo: true, importo: "200000", tassoAnnuo: "5", meseErogazione: currentYm(), rimborsoAScadenza: true },
      investitori: { attivo: true, lista: [
        { id: uid(), nome: "Mario Rossi", importo: "30000", meseVersamento: currentYm() },
        { id: uid(), nome: "Studio Bianchi Srl", importo: "20000", meseVersamento: currentYm() },
      ] },
    };
    setOperations((prev) => [...prev, demo]);
    setLiquidity((prev) => {
      const totale = totaleLiquidita(prev);
      const conti = totale > 0 ? prev.conti : [
        { id: uid(), nome: "Unicredit business", saldo: "420000" },
        { id: uid(), nome: "Unicredit deposito", saldo: "180000" },
      ];
      return { ...prev, conti };
    });

    const rentalDemo = emptyRental();
    rentalDemo.nome = "Bilocale Duomo";
    rentalDemo.indirizzo = "Milano, centro — bilocale arredato";
    rentalDemo.tipoAffitto = "breve";
    rentalDemo.meseInizioLocazione = addMonths(currentYm(), 2);
    rentalDemo.ricavi = { tariffaNotte: "130", occupazionePct: "68", canoneMensile: "" };
    rentalDemo.costiUnaTantum = [
      { id: uid(), categoria: "Acquisto", importo: "260000", data: todayStr() },
      { id: uid(), categoria: "Ristrutturazione / Lavori", importo: "35000", data: toDateStr(addMonths(currentYm(), 1)) },
      { id: uid(), categoria: "Arredo e allestimento", importo: "18000", data: toDateStr(addMonths(currentYm(), 2)) },
    ];
    rentalDemo.costiRicorrenti = [
      { id: uid(), categoria: "Gestione / Property management", importoMensile: "450", meseInizio: addMonths(currentYm(), 2), meseFine: "" },
      { id: uid(), categoria: "Utenze", importoMensile: "160", meseInizio: addMonths(currentYm(), 2), meseFine: "" },
      { id: uid(), categoria: "IMU", importoMensile: "90", meseInizio: currentYm(), meseFine: "" },
    ];
    rentalDemo.mutuo = { attivo: true, importo: "150000", tassoAnnuo: "3.4", durataMesi: "240", meseErogazione: currentYm() };
    setRentals((prev) => [...prev, rentalDemo]);

    setLiquidity((prev) => ({ ...prev, partite: [
      { id: uid(), tipo: "credito", controparte: "Acquirente Via Verdi", importo: "40000", data: toDateStr(addMonths(currentYm(), 4)), saldato: false, descrizione: "Saldo caparra" },
      { id: uid(), tipo: "debito", controparte: "Impresa edile", importo: "25000", data: toDateStr(addMonths(currentYm(), 1)), saldato: false, descrizione: "SAL lavori" },
    ], portafoglio: { titoli: [
      { id: uid(), nome: "Apple Inc.", ticker: "AAPL", quantita: "200", prezzoCarico: "150", prezzoAttuale: "210", valuta: "USD" },
      { id: uid(), nome: "iShares MSCI World ETF", ticker: "IWDA", quantita: "500", prezzoCarico: "80", prezzoAttuale: "95", valuta: "EUR" },
    ] } }));
  };

  // Esporta tutti i dati (operazioni, immobili, liquidità, bilancio) in un file JSON scaricabile.
  // Serve a spostare i dati tra copie diverse di questo artifact (es. da account personale a Team),
  // dato che lo storage condiviso è legato alla singola copia dell'artifact e non si sposta da solo.
  const exportData = () => {
    const payload = JSON.stringify({ liquidity, operations, rentals, bilancio }, null, 2);
    const blob = new Blob([payload], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `qg-real-estate-backup-${todayStr()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // Importa un file JSON precedentemente esportato con exportData, sovrascrivendo i dati correnti.
  // L'operazione è distruttiva e non annullabile, quindi valida il contenuto e chiede conferma
  // quando ci sono già dati in questa copia dell'artifact.
  const importData = (file) => {
    setImportError("");
    const reader = new FileReader();
    reader.onload = () => {
      let parsed;
      try {
        parsed = JSON.parse(reader.result);
      } catch (e) {
        setImportError("File non valido: non è un JSON leggibile.");
        return;
      }

      const haContenuto = parsed && typeof parsed === "object" &&
        (parsed.liquidity || Array.isArray(parsed.operations) || Array.isArray(parsed.rentals) || parsed.bilancio);
      if (!haContenuto) {
        setImportError("Questo file non sembra un backup di QG Real Estate: nessun dato riconosciuto.");
        return;
      }

      const datiPresenti = operations.length > 0 || rentals.length > 0 || (liquidity.conti || []).some((c) => Number(c.saldo) > 0);
      if (datiPresenti) {
        const riepilogo = [
          Array.isArray(parsed.operations) ? `${parsed.operations.length} operazioni` : null,
          Array.isArray(parsed.rentals) ? `${parsed.rentals.length} immobili a reddito` : null,
        ].filter(Boolean).join(", ");
        const ok = window.confirm(
          `Stai per sostituire TUTTI i dati di questa copia con il backup${riepilogo ? ` (${riepilogo})` : ""}.\n\n` +
          "I dati attuali andranno persi e l'operazione non è annullabile.\n" +
          "Se non l'hai già fatto, annulla ed esporta prima un backup di sicurezza.\n\nProcedere?"
        );
        if (!ok) { setImportError("Importazione annullata: i dati attuali sono intatti."); return; }
      }

      try {
        if (parsed.liquidity) setLiquidity(normalizeLiquidity(parsed.liquidity));
        if (parsed.operations) setOperations(parsed.operations.map(normalizeOperation));
        if (parsed.rentals) setRentals(parsed.rentals.map(normalizeRental));
        if (parsed.bilancio) setBilancio(normalizeBilancio(parsed.bilancio));
        setOpenOpId(null); setEditingOp(null); setOpenRentalId(null); setEditingRental(null);
        setTab("dashboard");
        setSaveStatus("Dati importati");
      } catch (e) {
        setImportError("Backup danneggiato: alcuni dati non sono leggibili.");
      }
    };
    reader.onerror = () => setImportError("Errore durante la lettura del file.");
    reader.readAsText(file);
  };

  const allInvestorNames = useMemo(() => {
    const names = new Set();
    operations.forEach((op) => (op.copertura?.investitori?.lista || []).forEach((inv) => { if (inv.nome) names.add(inv.nome); }));
    return Array.from(names);
  }, [operations]);

  const currentOp = openOpId ? operations.find((o) => o.id === openOpId) : null;
  const currentRental = openRentalId ? rentals.find((r) => r.id === openRentalId) : null;

  if (loading) {
    return (
      <div className="qg-root" style={{ display: "flex", alignItems: "center", justifyContent: "center", minHeight: 400 }}>
        <GlobalStyle />
        <div style={{ color: "var(--text-muted)", fontSize: 13 }}>Caricamento…</div>
      </div>
    );
  }

  return (
    <div className="qg-root" style={{ display: "flex", minHeight: 640, height: "100%", borderRadius: 14, overflow: "hidden", border: "1px solid var(--border)" }}>
      <GlobalStyle />

      {/* Sidebar */}
      <div style={{ width: 220, background: "var(--panel)", borderRight: "1px solid var(--border-soft)", display: "flex", flexDirection: "column", padding: "20px 12px", flexShrink: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "0 8px", marginBottom: 26 }}>
          <div style={{ width: 30, height: 30, borderRadius: 7, background: "var(--accent-soft)", display: "flex", alignItems: "center", justifyContent: "center", border: "1px solid var(--accent)" }}>
            <Building2 size={15} color="var(--accent)" />
          </div>
          <div>
            <div className="qg-display" style={{ fontSize: 15, fontWeight: 600, lineHeight: 1.1 }}>QG Real Estate</div>
            <div style={{ fontSize: 10, color: "var(--text-dim)", letterSpacing: "0.04em" }}>CONTROLLO LIQUIDITÀ</div>
          </div>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
          <NavItem icon={LayoutDashboard} label="Dashboard" active={tab === "dashboard" && !currentOp && !editingOp} onClick={() => { setTab("dashboard"); setOpenOpId(null); setEditingOp(null); setOpenRentalId(null); setEditingRental(null); }} />
          <NavItem icon={ListTree} label="Operazioni" active={tab === "operazioni" && !currentOp && !editingOp} onClick={() => { setTab("operazioni"); setOpenOpId(null); setEditingOp(null); setOpenRentalId(null); setEditingRental(null); }} count={operations.length} />
          <NavItem icon={KeyRound} label="Immobili a reddito" active={tab === "immobili" && !currentRental && !editingRental} onClick={() => { setTab("immobili"); setOpenOpId(null); setEditingOp(null); setOpenRentalId(null); setEditingRental(null); }} count={rentals.length} />
          <NavItem icon={Users} label="Investitori" active={tab === "investitori" && !currentOp && !editingOp} onClick={() => { setTab("investitori"); setOpenOpId(null); setEditingOp(null); setOpenRentalId(null); setEditingRental(null); }} />
          <NavItem icon={Scale} label="Bilancio" active={tab === "bilancio" && !currentOp && !editingOp} onClick={() => { setTab("bilancio"); setOpenOpId(null); setEditingOp(null); setOpenRentalId(null); setEditingRental(null); }} />
          <NavItem icon={Sparkles} label="Chiedi a Claude" active={tab === "askclaude" && !currentOp && !editingOp} onClick={() => { setTab("askclaude"); setOpenOpId(null); setEditingOp(null); setOpenRentalId(null); setEditingRental(null); }} />
          <NavItem icon={Receipt} label="Costi Fissi" active={tab === "costifissi" && !currentOp && !editingOp} onClick={() => { setTab("costifissi"); setOpenOpId(null); setEditingOp(null); setOpenRentalId(null); setEditingRental(null); }} />
          <NavItem icon={Settings2} label="Liquidità" active={tab === "liquidita" && !currentOp && !editingOp} onClick={() => { setTab("liquidita"); setOpenOpId(null); setEditingOp(null); setOpenRentalId(null); setEditingRental(null); }} />
        </div>

        <div style={{ marginTop: "auto", padding: "0 8px", display: "flex", flexDirection: "column", gap: 10 }}>
          {operations.length === 0 && rentals.length === 0 && (
            <button className="qg-btn qg-btn-ghost" style={{ fontSize: 11.5, justifyContent: "center" }} onClick={seedDemo}>Carica esempio</button>
          )}
          <div style={{ display: "flex", gap: 6 }}>
            <button
              className="qg-btn qg-btn-ghost"
              style={{ fontSize: 11, justifyContent: "center", flex: 1, padding: "6px 8px" }}
              onClick={exportData}
              title="Scarica un backup JSON di tutti i dati"
            >
              <Download size={13} /> Esporta
            </button>
            <button
              className="qg-btn qg-btn-ghost"
              style={{ fontSize: 11, justifyContent: "center", flex: 1, padding: "6px 8px" }}
              onClick={() => importInputRef.current?.click()}
              title="Carica un backup JSON esportato in precedenza"
            >
              <Upload size={13} /> Importa
            </button>
            <input
              ref={importInputRef}
              type="file"
              accept="application/json"
              style={{ display: "none" }}
              onChange={(e) => { const f = e.target.files?.[0]; if (f) importData(f); e.target.value = ""; }}
            />
          </div>
          {importError && (
            <div style={{ fontSize: 10.5, color: "var(--negative)", textAlign: "center" }}>{importError}</div>
          )}
          <div style={{ fontSize: 10.5, color: "var(--text-dim)", textAlign: "center" }}>{saveStatus || "Dati condivisi col team"}</div>
        </div>
      </div>

      {/* Main content */}
      <div className="qg-scroll" style={{ flex: 1, overflowY: "auto", background: "var(--bg)" }}>
        {editingOp ? (
          <OperationEditor
            operation={editingOp}
            liquidity={liquidity}
            allInvestorNames={allInvestorNames}
            onSave={saveOperation}
            onCancel={() => { if (operations.some((o) => o.id === editingOp.id)) { setOpenOpId(editingOp.id); } else { setTab("operazioni"); } setEditingOp(null); }}
            onDelete={operations.some((o) => o.id === editingOp.id) ? deleteOperation : null}
          />
        ) : currentOp ? (
          <OperationDetail
            op={currentOp}
            liquidity={liquidity}
            operations={operations}
            rentals={rentals}
            onEdit={() => setEditingOp(currentOp)}
            onBack={() => { setOpenOpId(null); setTab(returnTab); }}
          />
        ) : editingRental ? (
          <RentalEditor
            rental={editingRental}
            liquidity={liquidity}
            operations={operations}
            onSave={saveRental}
            onCancel={() => { if (rentals.some((r) => r.id === editingRental.id)) { setOpenRentalId(editingRental.id); } else { setTab("immobili"); } setEditingRental(null); }}
            onDelete={rentals.some((r) => r.id === editingRental.id) ? deleteRental : null}
          />
        ) : currentRental ? (
          <RentalDetail
            rental={currentRental}
            liquidity={liquidity}
            operations={operations}
            rentals={rentals}
            onEdit={() => setEditingRental(currentRental)}
            onBack={() => { setOpenRentalId(null); setTab("immobili"); }}
          />
        ) : tab === "dashboard" ? (
          <Dashboard liquidity={liquidity} operations={operations} rentals={rentals} onOpen={(id) => { setReturnTab("dashboard"); setOpenOpId(id); }} />
        ) : tab === "operazioni" ? (
          <OperationsList
            operations={operations}
            liquidity={liquidity}
            rentals={rentals}
            onOpen={(id) => { setReturnTab("operazioni"); setOpenOpId(id); }}
            onNew={() => setEditingOp(emptyOperation())}
            filterStato={filterStato}
            setFilterStato={setFilterStato}
          />
        ) : tab === "immobili" ? (
          <RentalsList
            rentals={rentals}
            operations={operations}
            onOpen={(id) => setOpenRentalId(id)}
            onNew={() => setEditingRental(emptyRental())}
          />
        ) : tab === "investitori" ? (
          <InvestorsView
            operations={operations}
            liquidity={liquidity}
            onOpenOperation={(id) => { setReturnTab("investitori"); setOpenOpId(id); }}
          />
        ) : tab === "costifissi" ? (
          <CostiFissiSettings liquidity={liquidity} setLiquidity={setLiquidity} />
        ) : tab === "bilancio" ? (
          <BilancioView bilancio={bilancio} setBilancio={setBilancio} />
        ) : tab === "askclaude" ? (
          <AskClaudeView liquidity={liquidity} operations={operations} rentals={rentals} bilancio={bilancio} />
        ) : (
          <LiquiditySettings liquidity={liquidity} setLiquidity={setLiquidity} />
        )}
      </div>
    </div>
  );
}
