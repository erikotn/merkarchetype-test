# Setup: online team-sessies aanzetten

Eenmalig ~10 minuten werk. Daarna kun je voor altijd online team-sessies hosten via de tool.

---

## Stap 1 — Google Sheet aanmaken

1. Open https://sheets.new (maakt direct een nieuwe lege Sheet aan)
2. Geef de Sheet een naam, bv. **"Merkarchetype Sessies"** (klik linksboven op "Untitled spreadsheet")
3. Laat de Sheet open staan

> ℹ️ De Sheet komt in jóuw Google Drive. Alleen jij hebt toegang. De tabbladen 'Sessions' en 'Responses' worden automatisch aangemaakt zodra de eerste sessie start.

---

## Stap 2 — Apps Script openen

In de Sheet:
- Klik in de bovenbalk op **Extensies** → **Apps Script**
- Er opent een nieuw tabblad met een leeg `Code.gs`-bestand

Verwijder de standaard inhoud (`function myFunction() {}`).

---

## Stap 3 — Code plakken

1. Open in deze repo: `apps-script/Code.gs` (https://github.com/erikotn/merkarchetype-test/blob/main/apps-script/Code.gs)
2. Klik rechtsboven op **Raw**
3. Selecteer alles (Cmd+A), kopieer (Cmd+C)
4. Plak in het Apps Script `Code.gs` tabblad
5. Klik op het 💾-icoon (Save) of Cmd+S

---

## Stap 4 — Deploy as Web App

1. Klik rechtsboven op **Deploy** → **New deployment**
2. Klik op het tandwiel naast "Select type" → kies **Web app**
3. Vul in:
   - **Description**: `Merkarchetype API v1`
   - **Execute as**: `Me (jouw@email.com)`
   - **Who has access**: `Anyone` ⚠️ (verplicht — zonder dit kunnen deelnemers niets indienen)
4. Klik **Deploy**
5. Eerste keer: er volgt een Google-toestemmingsdialoog
   - "Authorize access" → kies je account
   - Je ziet "Google hasn't verified this app" → klik **Advanced** → **Go to [project name] (unsafe)**
   - (Dit is jóuw eigen script — niet onveilig. Google waarschuwt voor alle ongecertificeerde Apps Scripts)
   - Klik **Allow**

6. Je krijgt een venster met **Web app URL**. Ziet er ongeveer zo uit:
   ```
   https://script.google.com/macros/s/AKfycbz.../exec
   ```
   **Kopieer die URL** (klik op het kopieer-icoon ernaast).

---

## Stap 5 — URL plakken in de tool

1. Open in deze repo: `index.html`
2. Zoek (Ctrl+F) naar `APPS_SCRIPT_URL`
3. Vervang de lege quotes:
   ```js
   const APPS_SCRIPT_URL = '';
   ```
   met de URL die je net kopieerde:
   ```js
   const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbz.../exec';
   ```
4. Sla op, commit en push naar `main`. GitHub Pages deployt automatisch binnen ~1 minuut.

---

## Stap 6 — Testen

1. Open https://erikotn.github.io/merkarchetype-test/
2. Klik op **Team online**
3. Geef de sessie een naam (bv. "Test")
4. Je krijgt twee links — kopieer de **deelnemers-link**
5. Open die deelnemers-link in een ander browservenster (incognito) → vul de test in → verzend
6. Ga terug naar je host-tabblad → klik op refresh — je ziet 1 deelnemer.

Werkt? Klaar.

---

## Wijzigingen aan Code.gs later

Als je het Apps Script later aanpast, moet je opnieuw deployen:
- **Deploy** → **Manage deployments** → klik op het 🖉-icoon naast je bestaande deployment
- Verander niets, klik op **Version: New version**, klik op **Deploy**
- De URL blijft hetzelfde (mits je de bestaande deployment update i.p.v. een nieuwe maakt)

---

## Wat als je de tool wilt resetten

Wil je alle sessies wissen? Open je Google Sheet, selecteer alle rijen behalve de eerste (header), rechtsklik → **Delete rows**. Doe dit in beide tabbladen.

Wil je de héle backend uitschakelen? Maak `APPS_SCRIPT_URL` weer leeg in `index.html`. De solo- en team-lokaal-modus blijven gewoon werken.
