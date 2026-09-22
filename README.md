# QG Real Estate — gestionale

App React per operazioni, liquidità e flussi di cassa di QG Real Estate, con dati su Supabase.
Contesto, architettura e regole di calcolo sono in `CLAUDE.md`.

## Setup (una volta sola)

### 1. Progetto Supabase

Crea un progetto su [supabase.com](https://supabase.com). Per i dati aziendali scegli una regione
UE (es. Frankfurt).

### 2. Database

Da **SQL Editor** esegui il contenuto di `supabase/migrations/20260922000000_init.sql`
(oppure, con la Supabase CLI collegata al progetto: `supabase db push`).

Poi autorizza i soci, sostituendo con le email reali:

```sql
insert into public.team_members (email, nome) values
  ('marco@qgrealestate.it', 'Marco'),
  ('luca@...', 'Luca'),
  ('alberto@...', 'Alberto');
```

### 3. Autenticazione

Si entra con **nome utente e password**: l'app non invia mai email. Chi scrive solo il nome
utente (`alberto`) accede all'account `alberto@qgrealestate.it`.

In **Authentication**:
- **Sign In / Providers → Email**: attivo, con **Confirm email** disattivato.
- Disattiva **Allow new users to sign up**: gli account li crea un amministratore.
- **Users → Add user → Create new user** per ciascun socio: email uguale a quella in
  `team_members`, password a scelta, **Auto Confirm User** attivo.

Password dimenticata: un amministratore la reimposta da **Users → … → Reset password**.

### 4. Funzione per "Chiedi a Claude" e quotazioni

Serve una API key da [console.anthropic.com](https://console.anthropic.com). Con la Supabase CLI:

```bash
supabase link --project-ref <ref-del-progetto>
supabase secrets set ANTHROPIC_API_KEY=sk-ant-...
supabase secrets set CLAUDE_MODEL=claude-sonnet-5   # opzionale
supabase functions deploy claude-proxy
```

Senza questo passo l'app funziona, ma "Chiedi a Claude", "Aggiorna quotazioni" e il fallback del
cambio USD/EUR restituiscono errore. Le chiamate sono a consumo sull'account Anthropic.

### 5. Avvio

```bash
cp .env.example .env.local   # inserisci URL e chiave anon (Project Settings → API)
npm install
npm run dev
```

### 6. Portare i dati dall'artifact

1. Nell'artifact attuale premi **Esporta** (barra laterale in basso): scarichi il backup JSON.
2. Nella nuova app, dopo il login, premi **Importa** e seleziona quel file.

Controlla poi che saldi e operazioni coincidano con l'artifact prima di smettere di usarlo.

## Pubblicazione

La build (`npm run build`) è un sito statico in `dist/`, pubblicabile su Vercel, Netlify o
Cloudflare Pages. Imposta lì le variabili `VITE_SUPABASE_URL` e `VITE_SUPABASE_ANON_KEY` e aggiungi
l'indirizzo pubblico ai Redirect URLs di Supabase.

## Recuperare una versione precedente

Ogni salvataggio conserva la versione precedente. Dal SQL Editor:

```sql
-- elenco delle versioni salvate
select id, version, archived_at from app_state_history
where key = 'qg-real-estate-data' order by archived_at desc limit 20;

-- ripristino di una versione (sostituisci l'id); la versione corrente finisce a sua volta nello storico
update app_state
   set data = (select data from app_state_history where id = <ID>),
       version = version + 1, updated_at = now()
 where key = 'qg-real-estate-data';
```

Dopo il ripristino, chi ha l'app aperta deve ricaricare la pagina.
