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

## Stap 7 — Admin-wachtwoord instellen

Hiermee kun je vanaf elk apparaat al je klantsessies terugvinden, ook nadat je je browsercache hebt gewist of een ander apparaat gebruikt.

In de Apps Script editor (waar je `Code.gs` net hebt geplakt):

1. **Bovenaan** zit een dropdown naast het ▶ Run-knopje. Klik erop en kies **`eenmaligWachtwoordZetten`**
2. Klik op **▶ Run**
3. Eerste keer: Google vraagt om toestemming → Allow
4. Een dialoog verschijnt in jouw Google Sheet (kan in een ander browsertabblad zitten — tab wisselen)
5. Type je wachtwoord, klik OK
6. Bevestiging: "Wachtwoord opgeslagen"

Geen redeploy nodig — direct actief.

> 💡 Wachtwoord later wijzigen? Run gewoon `eenmaligWachtwoordZetten` opnieuw. De oude wordt overschreven.

**Hoe gebruik je het in de tool:**
- Open https://erikotn.github.io/merkarchetype-test/
- Onderaan klikken op **⚙️ Beheer**
- Wachtwoord invullen → je ziet alle sessies met directe links

> ℹ️ Het wachtwoord blijft op jouw apparaat onthouden in localStorage tot je op "Uitloggen" klikt. Op een nieuw apparaat: typ 'm één keer.

---

## Stap 8 — Content-tabs aanmaken (vragen + archetypen bewerkbaar via Sheet)

Hiermee komen de vragen, archetype-omschrijvingen, scoring en eigen merkpersoonlijkheid-teksten in jouw Sheet te staan, zodat je ze daar makkelijk kunt aanpassen zonder code.

In de Apps Script editor (na het pasten van de nieuwe Code.gs):

1. Dropdown naast ▶ Run → kies **`eenmaligContentTabsAanmaken`**
2. Klik **▶ Run**
3. Bevestiging: er zijn vier nieuwe tabs in je Sheet:
   - **Archetypes** — naam, emoji, kleur, beschrijving, kernbelofte, communicatie, kernwaarden, voorbeeldmerken (12 rijen)
   - **Questions** — de 30 stellingen (left/right per vraag)
   - **Mappings** — punten per (vraag, antwoord, archetype). Antwoord 3 (neutraal) staat er niet in — dat is altijd 0 punten
   - **Combinations** — eigen merkpersoonlijkheid-tekst per top-2-combinatie (begint met 1 voorbeeld, vul aan over tijd)

**Aanpassen:** open de Sheet, ga naar de juiste tab, wijzig wat je wilt, save (gebeurt automatisch). De tool leest deze tabs live.

**Cache:** de tool cached de content 1 uur in de browser. Wil je je wijzigingen direct zien?
- Open de tool → ⚙️ Beheer → wachtwoord invullen → klik **🔃 Cache wissen** → refresh de pagina

**Nieuwe combination toevoegen:** open de Combinations-tab, voeg rij toe met:
- `primary`: archetype-key (bv. `hero`, `lover`, `creator` — de Engelstalige interne namen, zie kolom `key` in Archetypes-tab)
- `secondary`: tweede archetype-key
- `description`: jouw geschreven merkpersoonlijkheid-tekst

Wanneer iemand de test maakt en uitkomt op die combinatie, zien ze JOUW geschreven tekst i.p.v. niets.

> 💡 De tool zoekt bidirectioneel: als je `hero + sage` schrijft, werkt 'ie ook voor mensen die uitkomen op `sage + hero`. Je hoeft elk paar maar één keer te schrijven.

**Alle 132 voorgeschreven combinaties laden?** Run `vervangAlleCombinaties` (zelfde dropdown-route als `eenmaligContentTabsAanmaken`). Vervangt de hele Combinations-tab met 132 directionele teksten (12 archetypen × 11 partners — A primair + B secundair is een andere tekst dan B primair + A secundair). Bestaande aanpassingen gaan verloren — vandaar dat de functie eerst om bevestiging vraagt.

De 4 kolommen na het runnen:
- `primary` — primaire archetype-key
- `secondary` — secundaire archetype-key
- `type` — natuurlijke aanvulling / productieve spanning / niet eenvoudig
- `description` — de strategische tekst (de UI toont 'm)

> 💡 De type-kolom verschijnt in de tool als een gekleurde badge: groen voor natuurlijke aanvulling, oranje voor productieve spanning, rood voor niet eenvoudig.

---

## Stap 9 — Diagram uploaden (optioneel maar mooi)

Het intro-scherm toont een afbeelding van het merkarchetype-wiel. Upload het bestand `archetype-wheel.png` naar de root van je GitHub-repo:

1. Ga naar https://github.com/erikotn/merkarchetype-test
2. Klik **Add file** → **Upload files**
3. Sleep je PNG erin (zorg dat de bestandsnaam `archetype-wheel.png` is — hernoem 'm anders eerst lokaal)
4. Commit changes

Geen image? De tool verbergt 'm automatisch (geen kapotte placeholder).

---

## Wijzigingen aan Code.gs later

Als je het Apps Script later aanpast, moet je opnieuw deployen:
- **Deploy** → **Manage deployments** → klik op het 🖉-icoon naast je bestaande deployment
- Verander niets, klik op **Version: New version**, klik op **Deploy**
- De URL blijft hetzelfde (mits je de bestaande deployment update i.p.v. een nieuwe maakt)

---

## Waar staat mijn data?

Alles in jouw Google Sheet (in jouw Drive). Twee tabbladen:

- **Sessions**: id, name, host_token, status, created_at — één rij per sessie
- **Responses**: session_id, participant_name, answers (JSON-string), submitted_at — één rij per ingediend antwoord

Je kunt direct in de Sheet bekijken wie wat heeft ingediend. De `host_token` is je toegangscode per sessie — bewaar die zorgvuldig (of gebruik de admin-flow om hem altijd terug te vinden).

---

## Wat als je de tool wilt resetten

Wil je alle sessies wissen? Open je Google Sheet, selecteer alle rijen behalve de eerste (header), rechtsklik → **Delete rows**. Doe dit in beide tabbladen.

Wil je de héle backend uitschakelen? Maak `APPS_SCRIPT_URL` weer leeg in `index.html`. De solo- en team-lokaal-modus blijven gewoon werken.
