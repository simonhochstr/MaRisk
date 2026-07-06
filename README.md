# MaRisk-Änderungsanalyse

Werkzeug zur Auswertung der BaFin-Vergleichsversion des MaRisk-Rundschreibens. Aus dem PDF mit rot/grün/lila markierten Änderungen wird eine strukturierte Excel-Datei mit einer Zeile pro Textziffer.

## Zweck

Die BaFin veröffentlicht Konsultationen und Novellen der MaRisk als Vergleichs-PDF (gestrichen rot, neu rot unterstrichen, verschoben grün, Tz-Umnummerierungen lila). Das ist zum Lesen geeignet, aber nicht filterbar, sortierbar oder kommentierbar. Dieses Tool wandelt das PDF in eine **Excel-Arbeitsgrundlage** um, in der jede Textziffer eine Zeile ist, mit vollständiger Rich-Text-Darstellung der Änderungen.

Typische Anwendung: Impact-Assessment für die Umsetzung der Novelle in einer Bank oder Beratung.

## Anforderungen

- Python 3.10+
- `pymupdf` (fitz) zum PDF-Parsing
- `openpyxl` für die Excel-Ausgabe

Installation:

```sh
pip3 install --break-system-packages pymupdf openpyxl
```

## Eingabedateien

Im Projektverzeichnis liegen:

- **Vergleichs-PDF** (Pflicht) — das markierte BaFin-Vergleichsdokument. Dateiname in [analyze.py](analyze.py) unter `PDF`.
- **Neue Einzelfassung** (optional, für Spalte A) — die unmarkierte neue MaRisk-Fassung.
- **Alte Einzelfassung** (optional, für Spalte B) — die unmarkierte alte MaRisk-Fassung.

Die beiden Einzelfassungen dienen als **zuverlässige Referenz** für die Spalten A/B: Der Textkörper jeder Tz wird gegen sie abgeglichen, um die neue bzw. alte Referenz zu bestimmen — unabhängig von der teils unzuverlässigen Abschnitts-/Seitenstruktur der Vergleichsfassung. Tz, die in der neuen Einzelfassung vorhanden sind, in der Vergleichsfassung aber nicht sauber zugeordnet werden konnten, werden aus der Einzelfassung **ergänzt**. Zuordnung von Dateiname → Referenzen und → Layout-Profil in [analyze.py](analyze.py) unter `REFERENCES` bzw. `DOCUMENTS`.

### Layout-Profile

Verschiedene BaFin-PDF-Generationen haben leicht unterschiedliche Layouts (Spaltenlage, Tz-Marginalspalte, Kopf-/Fußzeilen, Überschriftsgrößen). Diese Werte sind in [marisk_parser.py](marisk_parser.py) unter `PROFILES` je Dokumenttyp hinterlegt (`konsultation`, `vergleichsfassung`). Mehrseitige Inhaltsverzeichnisse werden automatisch erkannt und übersprungen (`is_toc_page`).

## Verwendung

```sh
python3 analyze.py
```

Das Skript parst alle Seiten und erzeugt `MaRisk_Aenderungsanalyse_pro_Textziffer.xlsx`.

## Aufbau der Excel-Ausgabe

**Arbeitsblatt „Änderungen pro Tz" (Hauptblatt)** — pro Textziffer **zwei Zeilen** untereinander: zuerst der Normtext, dann die Erläuterung:

| Spalte | Inhalt |
|---|---|
| A Textziffer | neue Referenz (Abschnitt + Tz-Nummer), per Textabgleich gegen die **neue Einzelfassung** bestimmt; leer, wenn die Tz vollständig gestrichen wurde |
| B alte Referenz | alte Referenz, per Textabgleich gegen die **alte Einzelfassung** bestimmt; leer, wenn die Tz neu hinzugefügt wurde |
| C Inhaltstyp | `Textziffer` (Normtext, linke PDF-Spalte), `Erläuterung` (rechte PDF-Spalte) oder `Überschrift` |
| D Inhalt | Rich-Text des jeweiligen Inhalts mit Farbmarkierungen und Strike/Underline; Trennungs-Bindestriche am Zeilenende sind entfernt |
| E Änderungsart | unverändert / geändert / gestrichen / hinzugefügt / verschoben (bezogen auf den Zeileninhalt) |
| F Verschiebung | heuristischer Ziel- oder Herkunfts-Code (Volltext- oder Teiltext-Match) |
| G Unsicher | `Ja`, wenn Verschiebungs-Match unter 75 % Ähnlichkeit |
| H Anmerkungen | automatische Diff-Summary + Verschiebungs-Vermerke inkl. Ähnlichkeit in %; `Teilverschiebung` = nur Teile des Textkörpers erscheinen an anderer Stelle; `aus Neufassung ergänzt` = Tz aus der neuen Einzelfassung nachgetragen |

**Arbeitsblatt „Legende"** — Erklärung aller Spalten und Farbcodes.

**Arbeitsblatt „Umbenennungen"** — kompakte Lookup-Tabelle `alt → neu` aller erkannten Abschnitts- und Tz-Umnummerierungen. Praktisch für VLOOKUP aus internen Referenzlisten.

### Zeilenfärbung

Der Zeilenhintergrund richtet sich nach dem „stärkeren" Änderungsstatus der beiden Text-Spalten:

| Farbe | Bedeutung |
|---|---|
| weiß | unverändert |
| hellblau `#DEEBF7` | geändert |
| hellrot `#FCE4E4` | gestrichen |
| hellgrün `#E2EFDA` | hinzugefügt |
| hellgelb `#FFF2CC` | verschoben vorher |
| gelbgrün `#EDFADE` | verschoben nachher |
| helllila `#EDE7F6` | geänderte Überschrift |

## Arbeiten mit der Excel

- **Autofilter setzen** (Menü *Daten → Filter*) und in Spalte A/E nach Bedarf filtern.
- **Alle umgezogenen Tz sehen**: Spalte B nach „nicht leer" filtern.
- **Zu einer bestimmten Tz springen**: Strg/Cmd+F nach dem Tz-Code in Spalte A.
- **VLOOKUP aus eigenen Listen**: auf Spalte B (alte Referenz) aufsetzen.

## Dateien im Projekt

| Datei | Rolle |
|---|---|
| [analyze.py](analyze.py) | Haupt-Skript, erzeugt die Excel |
| [marisk_parser.py](marisk_parser.py) | PDF-Parser-Bibliothek (Zeichen, Farben, Absätze, Rich-Text) |
| `dl_kon_*_rs_marisk-*_vergleichsversion.pdf` | Eingabe-PDF der BaFin (nicht im Repo, lokal abzulegen) |
| `MaRisk_Aenderungsanalyse_pro_Textziffer.xlsx` | generierte Excel-Ausgabe (nicht im Repo, wird vom Skript erzeugt) |

## Technik in Kurzform

- PyMuPDF liest Zeichen mit Position, Farbe, Font.
- Strike/Underline werden **nicht** per Font-Flag erkannt, sondern über dünne farbige Rechtecke im PDF; Position mittig → Strike, Position an der Baseline → Underline.
- Farben: Rot = Inhaltsänderung, Grün = Verschiebung, **Lila** (`#5C2E91`) = Tz-Umnummerierung in der Marge.
- Spaltentrennung bei x = 360 pt (links = Textziffer, rechts = Erläuterung).
- Tz-Bodies werden über Seitenumbrüche hinweg zusammengehalten.
- Fett gesetzte Passagen bleiben Teil der umgebenden Textziffer und werden nicht als eigene Einträge abgetrennt.
- Fußnoten-Marker in der Marge werden erkannt und aus der Auswertung entfernt.
- Pure-deleted alte Tz-Nummern werden als `alt Tz. N` in eigener Zeile dargestellt, damit sie nicht mit umnummerierten neuen Tz gleicher Nummer kollidieren.

### 5-Pass-Verschiebungsanalyse (Spalten G / I)

Für jede gestrichene Tz werden alle hinzugefügten Tz in fünf Stufen abgeglichen. Normtext **und** Erläuterung fließen in den Vergleich ein:

| Pass | Vergleichsbasis | Schwellenwert |
|---|---|---|
| 1 – Volltext | Gesamter Textkörper beider Seiten (bis 900 Z.) — Vorfilter für Passes 2–5 | ≥ 20 % |
| 2 – Absätze | Jeder Absatz der Quell-Tz als Gleitfenster im Zieltext | ≥ 68 % |
| 3 – Positionsfenster | Anfangs-, Mittel- und Endblock (~45 % der Länge) als Fenster im Zieltext | ≥ 68 % |
| 4 – Rückwärts | Absätze und Fenster der Ziel-Tz gegen den Quelltext | ≥ 68 % |
| 5 – Satz-Fingerabdruck | Anteil gemeinsamer Sätze (Einzelsatz-Ähnlichkeit ≥ 82 %) | ≥ 55 % |

Passes 2–5 laufen nur für die Top-20-Kandidaten aus Pass 1 sowie alle mit Pass-1-Score ≥ 20 %. Die Ähnlichkeit wird als Prozentwert in Spalte I ausgewiesen; Teilverschiebungen (nur ein Abschnitt des Textkörpers erscheint woanders) werden explizit als `Teilverschiebung` gekennzeichnet.

## Bekannte Einschränkungen

- Verschmolzene Alt-/Neu-Spans im Fließtext (z. B. `gibtzeigt` = alt `gibt` + neu `zeigt`) werden im Rich-Text korrekt per Strike/Underline unterschieden, sehen im reinen Text aber zusammengeklebt aus.
- Die Verschiebungs-Heuristik (Spalten G/H) ist ein Vorschlag; Treffer mit Spalte H = `Ja` sollten manuell geprüft werden.
- Vollständig gestrichene und neu eingefügte Tz, die strukturell einer Umnummerierung entsprechen (Blocktausch), erscheinen als getrennte Gestrichen-/Hinzugefügt-Zeilen ohne Eintrag in Spalte B — diese Fälle fängt die Verschiebungs-Heuristik auf.
- Teilverschiebungen (ein Absatz oder Block einer Tz taucht in einer anderen Tz auf) werden über Passes 2–4 erkannt und in Spalte I als `Teilverschiebung` gekennzeichnet. Der Schwellenwert liegt bei 68 % Ähnlichkeit (vs. 55 % beim Volltext-Match).
- Die Einlese-Logik ist auf den konkreten Aufbau der BaFin-Vergleichs-PDFs (Spaltenlayout, Schriftarten, Farbpalette) zugeschnitten. Bei anderen PDFs können Schwellenwerte angepasst werden müssen.
