# ₿ BTC Tracker — Saylor Style (cloud)

Dashboard HTML interattiva per tracciare acquisti di Bitcoin: grafico "Saylor
style" con bolle dimensionate sulla quantità, linea del prezzo medio d'acquisto,
zoom/pan e P&L in tempo reale.

**Questa è la versione pubblica "shell": non contiene alcun dato personale.**
Gli acquisti vengono letti a runtime dal tuo Google Sheet, solo nel tuo browser,
tramite un token segreto. Chi apre il sito senza il token non vede nulla.

## Come si usa
1. Apri la dashboard (GitHub Pages) — vedi solo l'interfaccia, senza dati.
2. Premi **☁ Collega…**, incolla l'URL del tuo Web App Apps Script (`…/exec`) e il tuo token.
3. Premi **⎓ Carica**: i tuoi acquisti vengono letti dal foglio.
4. URL e token restano salvati **solo nel tuo browser** (`localStorage`), mai nel sito.

## Backend
Vedi [`apps-script/Code.gs`](apps-script/Code.gs): crea il foglio pulito `REGISTRO`
e fornisce l'API (lettura via JSONP, scrittura via POST). Istruzioni di deploy nel file.

## Privacy
Il repository è pubblico ma **non contiene acquisti, importi o saldi**. I dati
vivono esclusivamente nel tuo Google Sheet, protetti dal token.
