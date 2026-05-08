# Merkarchetype-test

Welk merkarchetype past bij jouw merk? Test op basis van het 12-archetype-model (de Held, de Onschuldige, de Magiër, etc.) — handig als startpunt voor een positionerings- of merkstrategie-gesprek.

**👉 [Open de test](https://erikotn.github.io/merkarchetype-test/)**

## Drie modes

- **👤 Solo** — vul individueel in, krijg direct je top-3 archetype-mix
- **👥 Team — lokaal** — meerdere mensen om de tafel, één apparaat. Iedereen vult na elkaar in. Geen internet nodig
- **🌐 Team — online** — iedereen op eigen apparaat, tegelijk. Antwoorden landen in een Google Sheet die jij beheert. Vereist eenmalig setup — zie [SETUP.md](SETUP.md)

## Features

- 30 stellingen op een 5-puntsschaal
- Top-3 archetypen met percentages, kleuren, emoji en kerneigenschappen
- Nederlandse merkvoorbeelden per archetype (Campina, ANWB, Coolblue, Rituals, etc.)
- Auto-save in localStorage — refresh = je voortgang blijft
- Download als PDF (via je browser)
- Deelbare link voor solo-resultaten (antwoorden encoded in URL)
- Online sessies kunnen door de host gesloten worden — daarna geen nieuwe inzendingen meer
- Resultaten zijn alleen zichtbaar voor wie het host-token heeft

## Voor wie

Marketeers, ondernemers en strategen die een eerste richting zoeken voor merkpositionering, en bureaus die hem gebruiken als gespreks-opener of als gestructureerde input voor een merksessie.

## Tech

- 1 HTML-bestand (React via CDN, geen build-stap)
- Backend voor online team-modus: Google Apps Script + Google Sheet (zie [SETUP.md](SETUP.md))
