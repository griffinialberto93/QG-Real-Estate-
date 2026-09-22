# QG Real Estate — gestionale operazioni e liquidità

Tool interno di QG Real Estate srl (investimento e ristrutturazione immobiliare, Milano) per
monitorare operazioni, liquidità, flussi di cassa e sostenibilità, più il ramo Smart Rent Solution
(gestione affitti brevi). Nato come artifact di Claude, portato qui su Vite + React + Supabase.

Utenti: i tre soci (Marco, Luca, Alberto). Lingua dell'interfaccia: **italiano**.

## Comandi

```bash
npm install
npm run dev      # sviluppo su http://localhost:5173
npm run build    # build di produzione in dist/
```

Serve un file `.env.local` (vedi `.env.example`). Setup completo di Supabase nel `README.md`.

## Struttura

```
src/
  App.jsx            # l'intera app (~4.400 righe): motore di calcolo in alto, componenti UI sotto
  AuthGate.jsx       # login nome utente + password, pulsante Esci
  main.jsx
  lib/
    supabase.js      # client Supabase
    storage.js       # sostituto di window.storage dell'artifact (get/set su tabella app_state)
    claude.js        # chiamate a Claude tramite la Edge Function claude-proxy
supabase/
  migrations/        # schema: team_members, app_state, app_state_history, save_app_state()
  functions/claude-proxy/   # proxy verso l'Anthropic API (la chiave resta sul server)
```

### Come funziona la persistenza

- Tutto il workspace (liquidità, operazioni, immobili a reddito, bilancio) è **un unico documento
  JSON** nella riga `app_state` con `key = 'qg-real-estate-data'`. È lo stesso formato del backup
  "Esporta/Importa" della barra laterale: i dati dell'artifact si portano qui con Importa.
- Salvataggio automatico con debounce di 600 ms dopo ogni modifica.
- **Controllo di versione**: `save_app_state(key, data, expected_version)` rifiuta il salvataggio
  con `CONFLICT` se un altro socio ha salvato nel frattempo. L'app mostra "Non salvato: dati
  modificati da un altro utente. Ricarica la pagina." Non esiste merge automatico.
- Se il caricamento iniziale fallisce, il salvataggio resta disabilitato (uno stato vuoto non deve
  mai sovrascrivere i dati reali). Non rimuovere questa protezione.
- Ogni salvataggio che cambia i dati archivia la versione precedente in `app_state_history`.
- Accesso: login con nome utente e password (nessuna email inviata dall'app); leggono e scrivono
  solo gli utenti autenticati con email in `team_members` (policy RLS). Nessuna delete dal client.

### Funzioni esterne

- **Chiedi a Claude**, **Aggiorna quotazioni** titoli e fallback del **cambio USD/EUR** passano da
  `callClaude()` → Edge Function `claude-proxy`. Il modello lo decide il server (`CLAUDE_MODEL`),
  non il client. Non mettere mai la API key Anthropic nel frontend.
- Il cambio USD/EUR usa prima Frankfurter (dati BCE, CORS aperto) direttamente dal browser.
- Import estratto conto Unicredit: CSV ed Excel (SheetJS), rilevamento automatico delle colonne.

## Regole di business (vincolanti)

Fonte: `Regole business - QG Real Estate.docx` nel progetto Claude, più le decisioni prese durante
lo sviluppo. Se una modifica le contraddice, fermati e chiedi.

**Stati delle operazioni**
- Pipeline: valutata, non acquisita. In corso: acquisita. Completata: venduta. Abbandonata: fuori perimetro.
- Completate e abbandonate **non** entrano né nella proiezione né nello scadenzario spese.

**Flussi di cassa**
- Stime per mese aggregato. Orizzonti di proiezione standard: 12 / 24 / 36 / 60 mesi.
- I costi fissi di struttura sono aziendali, non allocati sulle singole operazioni.
- **Costi pieni + capitale in entrata**: il capitale di investitori e banca transita dal conto QG.
  I costi di un'operazione escono per intero; versamenti dei soci (`meseVersamento`) ed erogazione
  del prestito (`meseErogazione`) entrano come entrate nel loro mese. Non ridurre i costi alla quota QG.
- **Data di riconciliazione** (`liquidity.dataRiferimento`): tutto ciò che è datato fino a quella data
  è già dentro i saldi dei conti e non si proietta di nuovo. Vale per movimenti extra, costi una
  tantum, mesi dei costi ricorrenti (in base al `giornoAddebito`, vuoto = fine mese), rate ed
  entrate di capitale. Eccezione: debiti/crediti usano il loro flag `saldato`.
- **Rischio noto di doppio conteggio**: capitale già incassato ma datato dopo la riconciliazione
  viene proiettato come entrata futura. L'editor mostra un avviso; il pannello "Da dove viene questo
  numero" in Dashboard elenca ogni voce di ogni mese. Ogni `addFlow` deve portare un'etichetta leggibile.
- **Denaro di terzi**: ogni conto ha `vincolato` ("di cui di terzi", es. quota proprietari Smart Rent).
  È una passività: la proiezione parte da `totaleDisponibile = saldo − vincolato`.
- Il portafoglio titoli è una riserva liquidabile ma **non** entra nella proiezione di cassa.

**Coperture e investitori**
- Fonti: liquidità QG, prestito bancario, capitale investitori (rendimento a exit, non periodico).
- Flipping: prestito ponte, soli interessi, rimborso alla vendita. A reddito: mutuo con
  ammortamento francese ed eventuale pre-ammortamento (soli interessi).
- Utile investitore = quota del margine netto (dopo IRES + IRAP) proporzionale al capitale versato.
  Mostrare sempre utile lordo, imposta sostitutiva (26%) e utile netto.
- Frazionamento: più vendite con date proprie; rimborsi, imposte e liquidazione soci si ripartiscono
  in proporzione all'importo di ogni unità.

**Immobili a reddito**
- L'acquisto passa dalla sezione Operazioni con `destinazione = "reddito"`. La scheda in
  "Immobili a reddito" si collega con `operazioneId` ed **eredita** costi d'acquisto e mutuo senza
  duplicarli; lì si gestiscono solo ricavi e costi di gestione.

**Bilancio gestionale**
- Due rami: Lato Operazioni e Smart Rent Solution, inserimento manuale mensile.
- La **restituzione capitale** abbatte solo l'utile, mai il fatturato (il fatturato resta lordo).

**Test di sostenibilità** (in quest'ordine): 1) è in profitto? se no è sconsigliata e ci si ferma;
2) la liquidità QG, con tutte le operazioni attive insieme, resta sopra la soglia minima?
3) per quanti mesi regge senza vendere (autonomia)?

## Regole di lavoro

- Prima di presentare una modifica al motore di calcolo, **verificala con un test numerico**
  (scenario con numeri attesi calcolati a mano). Questa è una regola esplicita del committente.
- Mai stimare un dato mancante: chiedilo.
- Testi dell'interfaccia in italiano. Commenti nel codice in inglese o italiano, come quelli esistenti.
- Normalizzazione: ogni nuovo campo va aggiunto sia agli `empty*()` sia ai `normalize*()`, in modo
  che i backup vecchi continuino a caricarsi (retrocompatibilità obbligatoria).
- I backup JSON contengono saldi e nomi degli investitori: non committarli (sono in `.gitignore`).

## Punti aperti (da decidere con Marco)

1. **Pipeline nella proiezione**: il documento delle regole dice che la pipeline "impatta la
   proiezione solo in scenario"; il codice oggi la include nella proiezione base insieme a "In corso".
2. **Soglia minima**: le regole fissano € 10.000,00 (capitale sociale versato); il valore di default
   nell'app è 0 e va impostato in Liquidità.
3. **Formato importi**: le regole chiedono `€ 250.000,00`; `fmtEUR` oggi mostra `250.000 €` senza decimali.
4. Recharts 2.x è in fine supporto; valutare l'aggiornamento a 3.x.

## Refactoring suggerito (non ancora fatto)

`App.jsx` è un unico file perché nasceva come artifact. Ordine consigliato:
1. Estrarre il motore di calcolo puro (da `uid`/formattatori fino a `getInvestorsSummary`,
   `computeGlobalProjection`, `bilancioAnnuale`, parser bancari) in `src/engine/` **e scrivere test
   Vitest** sugli scenari numerici prima di toccare altro.
2. Poi spezzare i componenti UI (Dashboard, OperationEditor, RentalEditor, LiquiditySettings, …).
3. Code splitting per ridurre il bundle (oggi ~1,3 MB non compresso).
