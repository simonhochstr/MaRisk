# MaRisk-Änderungsanalyse

Werkzeug zur Auswertung der BaFin-Vergleichsversion des MaRisk-Rundschreibens. Aus dem PDF mit rot/grün/lila markierten Änderungen wird eine strukturierte Excel-Datei mit einer Zeile pro Textziffer.

## Zweck

Die BaFin veröffentlicht Konsultationen und Novellen der MaRisk als Vergleichs-PDF (gestrichen rot, neu rot unterstrichen, verschoben grün, Tz-Umnummerierungen lila). Das ist zum Lesen geeignet, aber nicht filterbar, sortierbar oder kommentierbar. Dieses Tool wandelt die 148 PDF-Seiten in eine **Excel-Arbeitsgrundlage** um, in der jede Textziffer eine Zeile ist, mit vollständiger Rich-Text-Darstellung der Änderungen.

Typische Anwendung: Impact-Assessment für die Umsetzung der Novelle in einer Bank oder Beratung.

## Anforderungen

- Python 3.10+
- `pymupdf` (fitz) zum PDF-Parsing
- `openpyxl` für die Excel-Ausgabe

Installation:

```sh
pip3 install --break-system-packages pymupdf openpyxl
```

## Eingabedatei

Die Vergleichs-PDF der BaFin muss im Projektverzeichnis liegen:

```
dl_kon_02_2026_rs_marisk-novelle_vergleichsversion.pdf
```

Der Dateiname ist in [analyze.py](analyze.py) unter der Konstante `PDF` festgelegt.

## Verwendung

```sh
python3 analyze.py
```

Das Skript parst alle Seiten und erzeugt `MaRisk_Aenderungsanalyse_pro_Textziffer.xlsx`.

## Aufbau der Excel-Ausgabe

**Arbeitsblatt „Änderungen pro Tz" (Hauptblatt)** — eine Zeile pro Textziffer:

| Spalte | Inhalt |
|---|---|
| A Textziffer | neue Bezeichnung, z. B. `AT 4.4.2 Tz. 5`; bei vollständig gestrichenen alten Tz: `AT 4.4.2 alt Tz. 4` |
| B alte Referenz | alter Pfad vor einer Umbenennung oder Umnummerierung; leer wenn unverändert |
| C Normtext | Rich-Text des linken Spaltentexts der Tz, mit Farbmarkierungen und Strike/Underline |
| D Erläuterung | Rich-Text aller zur Tz gehörenden rechten Spalten-Absätze |
| E Änderungsart Normtext | unverändert / geändert / gestrichen / hinzugefügt / verschoben |
| F Änderungsart Erläuterung | dasselbe Schema für die Erläuterung |
| G Verschiebung | heuristischer Ziel- oder Herkunfts-Code |
| H Unsicher | `Ja`, wenn Verschiebungs-Match unter 75 % Ähnlichkeit |
| I Anmerkungen | automatische Diff-Summary (Wortzahlen, Umformulierungs-Hinweis) + Verschiebungs-Vermerke |

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
| [dl_kon_02_2026_rs_marisk-novelle_vergleichsversion.pdf](dl_kon_02_2026_rs_marisk-novelle_vergleichsversion.pdf) | Eingabe-PDF der BaFin |
| [MaRisk_Aenderungsanalyse_pro_Textziffer.xlsx](MaRisk_Aenderungsanalyse_pro_Textziffer.xlsx) | generierte Excel-Ausgabe |

## Technik in Kurzform

- PyMuPDF liest Zeichen mit Position, Farbe, Font.
- Strike/Underline werden **nicht** per Font-Flag erkannt, sondern über dünne farbige Rechtecke im PDF; Position mittig → Strike, Position an der Baseline → Underline.
- Farben: Rot = Inhaltsänderung, Grün = Verschiebung, **Lila** (`#5C2E91`) = Tz-Umnummerierung in der Marge.
- Spaltentrennung bei x = 360 pt (links = Textziffer, rechts = Erläuterung).
- Tz-Bodies werden über Seitenumbrüche hinweg zusammengehalten.
- Fett gesetzte Passagen bleiben Teil der umgebenden Textziffer und werden nicht als eigene Einträge abgetrennt.
- Fußnoten-Marker in der Marge werden erkannt und aus der Auswertung entfernt.
- Pure-deleted alte Tz-Nummern werden als `alt Tz. N` in eigener Zeile dargestellt, damit sie nicht mit umnummerierten neuen Tz gleicher Nummer kollidieren.

## Bekannte Einschränkungen

- Verschmolzene Alt-/Neu-Spans im Fließtext (z. B. `gibtzeigt` = alt `gibt` + neu `zeigt`) werden im Rich-Text korrekt per Strike/Underline unterschieden, sehen im reinen Text aber zusammengeklebt aus.
- Die Verschiebungs-Heuristik (Spalten G/H) ist ein Vorschlag; Treffer mit Spalte H = `Ja` sollten manuell geprüft werden.
- Vollständig gestrichene und neu eingefügte Tz, die strukturell einer Umnummerierung entsprechen (Blocktausch), erscheinen als getrennte Gestrichen-/Hinzugefügt-Zeilen ohne Eintrag in Spalte B — diese Fälle fängt die Verschiebungs-Heuristik auf.
- Die Einlese-Logik ist auf den konkreten Aufbau der BaFin-Vergleichs-PDFs (Spaltenlayout, Schriftarten, Farbpalette) zugeschnitten. Bei anderen PDFs können Schwellenwerte angepasst werden müssen.
