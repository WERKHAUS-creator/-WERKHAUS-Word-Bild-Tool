# WERKHAUS Word & Bild

Office-Add-in für Microsoft Word zur Bildverwaltung, Beschriftung und Bilddokumentation.

## Zweck der Projektdatei

Die Datei `werkhaus_bilddaten.json` ist die zentrale Projektdatenbank pro Bildordner.
Sie speichert den aktuellen Projektstand und wird beim Arbeiten mit einem Bildordner
automatisch mitgeführt, wenn ein schreibbarer Ordnerhandle verfügbar ist.

## Verhalten beim Laden eines Bildordners

- Wenn im Ordner eine `werkhaus_bilddaten.json` vorhanden ist, wird sie automatisch erkannt und geladen.
- Die Bilder aus dem Ordner werden mit den Daten aus der Projektdatei zusammengeführt.
- Reihenfolge, Captions, Auswahl, Sichtbarkeit, Positionen, Bildnummern, Analyseprofile und Analyseergebnisse werden übernommen.
- Unbekannte Zusatzfelder bleiben im Projekt nach Möglichkeit erhalten.
- Wenn keine Projektdatei vorhanden ist, wird intern sofort eine neue Projektstruktur vorbereitet.

## Legacy-Unterstützung

- Alte Projektdateien `bilddaten.json` werden weiterhin erkannt.
- Legacy-Daten werden in das neue Projektformat übernommen.
- Der Benutzer erhält einen Hinweis, wenn eine Legacy-Datei geladen wurde.

## Roundtrip-Verhalten

Beim Export und erneuten Laden bleiben erhalten:

- Captions
- Reihenfolge
- Auswahl
- Sichtbarkeit
- Analyseprofile
- Analyseergebnisse
- unbekannte Felder, soweit technisch sinnvoll

Wichtige Regel:

- `images[].caption` ist die finale, verwendete Beschriftung.
- `analysis.*.result.suggested_caption` ist nur ein KI-Vorschlag.
- Benutzerbeschriftungen werden nicht ungefragt durch KI-Vorschläge überschrieben.

## Auto-Save und Fallback

Wenn der Ordner über einen schreibbaren Handle verfügbar ist:

- Änderungen werden mit kurzer Verzögerung automatisch gespeichert.
- Es wird `werkhaus_bilddaten.json` direkt im Ordner aktualisiert.

Wenn kein direkter Schreibzugriff verfügbar ist:

- Der Projektzustand bleibt intern aktuell.
- Der Speichern-Button exportiert die aktuelle Projekt-JSON als Fallback.
- Der Status weist auf den nötigen Export hin.

Wichtige Einschränkung:

- Auto-Save wurde in der Entwicklung vorbereitet und lokal geprüft.
- Ein vollständiger Live-End-to-End-Test im echten Word-/Office-Host sollte zusätzlich manuell geprüft werden.

## Bekannte Risiken

- Schreibbarkeit hängt vom jeweiligen Office-/Browser-Host ab.
- Absolute Pfade sind nur lokal aussagekräftig.
- Auto-Save sollte im echten Host einmal manuell geprüft werden.
- Fehlende Bilder werden gemeldet, aber nicht gelöscht.
- Die Word-Ausgabe bleibt von der Projektdatei-Logik getrennt.

## Wichtige Befehle

```bash
npm install
npm run dev-server
npm run start
npm run build
npm run validate:dev
npm run validate:prod
npm run check:manifests
npm run lint
```

## Struktur

- `src/` ist die Quelle.
- `assets/` enthält die eingebundenen Icons und Logos.
- `manifest.dev.xml` ist das lokale Manifest.
- `manifest.xml` ist das Produktionsmanifest.
- `dist/` ist nur Build-Output und wird bei Bedarf neu erzeugt.

## Deployment

- Lokale Entwicklung nutzt `https://localhost:3001`.
- Produktion nutzt `https://tool2.wh-sv.de`.
- GitHub Actions baut die App und deployt den erzeugten `dist/`-Ordner.
- `dist/manifest.xml` ist die veröffentlichte Manifest-Datei.

## Nächste Empfehlung

1. Zuerst den Live-Test im echten Word-/Office-Host mit echtem Ordnerhandle durchführen.
2. Danach gegebenenfalls kleine UI-/Status-Feinschliffe ergänzen.
3. Danach erst Release-Vorbereitung, Azure-Veröffentlichung und Admin-Center-Umstellung angehen.
