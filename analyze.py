"""
MaRisk-Vergleichsversion → Excel-Änderungsanalyse, aggregiert pro Textziffer.

Hauptskript für die Auswertung. Die PDF-Parser-Helfer
(load_page, group_paragraphs, split_footnote_from_tz, …) liegen in
marisk_parser.py. Dieses Skript baut darauf die Tz-Zeilen, fasst pro
Textziffer zusammen und schreibt die Excel-Datei.

Spaltenschema:
  A Textziffer (neu)       B alte Referenz     C Normtext
  D Erläuterung            E Änderungsart Norm F Änderungsart Erl.
  G Verschiebung           H Unsicher          I Anmerkungen
"""
import fitz
from difflib import SequenceMatcher
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font

from marisk_parser import (
    load_page, group_paragraphs, plain_text, segs_to_rich,
    classify_entry, SECTION_RE, _postprocess_xlsx,
    split_footnote_from_tz, diff_summary,
)

PDF = "dl_kon_02_2026_rs_marisk-novelle_vergleichsversion.pdf"
OUT = "MaRisk_Aenderungsanalyse_pro_Textziffer.xlsx"


def build_tz_rows(paragraphs):
    """Collapse paragraphs into one row per (section, Tz).

    Yields dicts with keys:
      label, section, norm_segs, expl_segs, notes, kind
    `kind` is either 'section_header' (pure heading row) or 'tz'.
    """
    current_section = ""
    current_section_old = ""   # old section code if renamed, else ""
    current_tz = None
    current_tz_old = None      # old Tz number if renumbered, else None
    last_new_tz = None         # integer counter of the last real "new" Tz number;
                               # drives the expected-next split context but is
                               # NOT touched by purely deleted old Tz rows.
    current_tz_is_alt = False  # True when the current row is a purely deleted
                               # old Tz (label "alt Tz. N"); suppresses collisions
                               # with the new numbering.
    tz_norm = []
    tz_expl = []
    tz_notes = []
    rows = []
    drop_next_footnote_body = False
    renames = []               # list of (old_code, new_code) for the Umbenennungen sheet

    def flush_tz():
        if current_tz is None and not tz_norm and not tz_expl:
            return
        if current_tz_is_alt:
            # Purely deleted old Tz – label with "alt Tz. N" so it does
            # not collide with the new Tz of the same number.
            label = (f"{current_section} alt Tz. {current_tz}"
                     if current_tz else current_section)
            old_ref = ""
        else:
            label = (f"{current_section} Tz. {current_tz}"
                     if current_tz else current_section)
            old_section_val = current_section_old or current_section
            old_tz_val      = current_tz_old or current_tz
            if current_section_old or current_tz_old:
                old_ref = (f"{old_section_val} Tz. {old_tz_val}"
                           if old_tz_val else old_section_val)
            else:
                old_ref = ""
        rows.append({
            "kind": "tz",
            "label": label,
            "old_ref": old_ref,
            "section": current_section,
            "tz": current_tz,
            "norm_segs": list(tz_norm),
            "expl_segs": list(tz_expl),
            "notes": list(tz_notes),
        })

    for p in paragraphs:
        txt = plain_text(p["segs"]).strip()
        if not txt:
            continue
        kind = p["kind"]
        col  = p["col"]

        if kind in ("section", "subsection"):
            flush_tz()
            tz_norm.clear(); tz_expl.clear(); tz_notes.clear()
            current_tz = None
            current_tz_old = None
            current_tz_is_alt = False
            last_new_tz = None
            drop_next_footnote_body = False
            # Extract new and old section codes.
            kept = plain_text(p["segs"],
                              fmt_filter={"unchanged", "added", "moved_to"}).strip()
            old  = plain_text(p["segs"],
                              fmt_filter={"unchanged", "deleted", "moved_from"}).strip()
            m_new = SECTION_RE.match(kept or txt)
            m_old = SECTION_RE.match(old)
            if m_new:
                current_section = m_new.group(0).replace("  ", " ").strip()
            current_section_old = ""
            if m_new and m_old:
                new_code = m_new.group(0).replace("  ", " ").strip()
                old_code = m_old.group(0).replace("  ", " ").strip()
                if new_code != old_code:
                    current_section_old = old_code
                    renames.append((old_code, new_code))
            rows.append({
                "kind": "section_header",
                "label": current_section or txt[:40],
                "old_ref": current_section_old,
                "section": current_section,
                "tz": None,
                "norm_segs": list(p["segs"]),
                "expl_segs": [],
                "notes": [],
            })
            continue

        if kind == "number":
            fmts = {s["fmt"] for s in p["segs"] if s["text"].strip()}
            new_digits = plain_text(p["segs"],
                                    fmt_filter={"unchanged", "added", "moved_to"}).strip()
            old_digits = plain_text(p["segs"],
                                    fmt_filter={"unchanged", "deleted", "moved_from"}).strip()

            if fmts == {"deleted"}:
                # Fully deleted old Tz — create a separate "alt" row,
                # keep last_new_tz untouched so the next real Tz still
                # sees the correct expected-next context.
                flush_tz()
                tz_norm.clear(); tz_expl.clear(); tz_notes.clear()
                current_tz = old_digits or txt
                current_tz_old = None
                current_tz_is_alt = True
                continue

            # Normal / added / renumbered case.
            tz_str, _ = split_footnote_from_tz(new_digits or txt, last_new_tz)
            if tz_str is None:
                drop_next_footnote_body = True
                continue
            old_tz_str = None
            if old_digits and old_digits != new_digits and old_digits.isdigit():
                old_tz_str = old_digits
            flush_tz()
            tz_norm.clear(); tz_expl.clear(); tz_notes.clear()
            current_tz = tz_str
            current_tz_old = old_tz_str
            current_tz_is_alt = False
            try:
                last_new_tz = int(tz_str)
            except ValueError:
                pass
            if old_tz_str:
                renames.append((f"{current_section} Tz. {old_tz_str}",
                                f"{current_section} Tz. {tz_str}"))
            continue

        if drop_next_footnote_body and col == "L" and kind == "body":
            drop_next_footnote_body = False
            continue
        drop_next_footnote_body = False

        # Body paragraph → append to current Tz buckets.
        target = tz_norm if col == "L" else tz_expl
        if target:
            target.append({"text": "\n", "fmt": "unchanged",
                           "flags": 0, "size": 10, "font": "Calibri"})
        target.extend(p["segs"])

    flush_tz()
    return rows, renames


def find_tz_moves(rows):
    """Simple heuristic: match gestrichene Tz-Normtexte zu neu hinzugefügten."""
    def del_text(r):
        return plain_text(r["norm_segs"],
                          fmt_filter={"deleted", "moved_from"}).strip()
    def add_text(r):
        return plain_text(r["norm_segs"],
                          fmt_filter={"added", "moved_to"}).strip()

    dels = [(i, del_text(r)) for i, r in enumerate(rows) if r["kind"] == "tz"]
    dels = [(i, t) for i, t in dels if len(t) >= 40]
    adds = [(i, add_text(r)) for i, r in enumerate(rows) if r["kind"] == "tz"]
    adds = [(i, t) for i, t in adds if len(t) >= 40]

    for i, dt in dels:
        best = 0.0
        bj = None
        for j, at in adds:
            if j == i:
                continue
            if abs(len(at) - len(dt)) > max(len(dt), len(at)) * 0.8:
                continue
            ratio = SequenceMatcher(None, dt[:500], at[:500]).ratio()
            if ratio > best:
                best = ratio
                bj = j
        if bj is not None and best >= 0.55:
            rows[i].setdefault("G", rows[bj]["label"])
            rows[i]["notes"].append(
                f"mögliche Verschiebung nach {rows[bj]['label']} (Ähnlichkeit {best:.2f})"
            )
            rows[bj].setdefault("G", rows[i]["label"])
            rows[bj]["notes"].append(
                f"mögliche Herkunft aus {rows[i]['label']} (Ähnlichkeit {best:.2f})"
            )
            if best < 0.75:
                rows[i]["uncertain"] = True
                rows[bj]["uncertain"] = True


def write_excel(rows, renames=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Änderungen pro Tz"
    headers = ["Textziffer", "alte Referenz", "Normtext", "Erläuterung",
               "Änderungsart Normtext", "Änderungsart Erläuterung",
               "Verschiebung", "Unsicher", "Anmerkungen"]
    ws.append(headers)
    header_font = Font(bold=True)
    for c in ws[1]:
        c.font = header_font
        c.alignment = Alignment(horizontal="left", vertical="top")

    fills = {
        "unverändert":       PatternFill("solid", fgColor="FFFFFFFF"),
        "geändert":          PatternFill("solid", fgColor="FFDEEBF7"),
        "gestrichen":        PatternFill("solid", fgColor="FFFCE4E4"),
        "hinzugefügt":       PatternFill("solid", fgColor="FFE2EFDA"),
        "verschoben vorher": PatternFill("solid", fgColor="FFFFF2CC"),
        "verschoben nachher":PatternFill("solid", fgColor="FFEDFADE"),
    }
    heading_fill = PatternFill("solid", fgColor="FFEDE7F6")
    # For the overall row colour use the "strongest" change of the two
    # columns — geändert > gestrichen/hinzugefügt > verschoben > unverändert.
    severity = {
        "unverändert": 0, "verschoben nachher": 1, "verschoben vorher": 1,
        "hinzugefügt": 2, "gestrichen": 2, "geändert": 3,
    }

    for r in rows:
        norm_rich = segs_to_rich(r["norm_segs"])
        expl_rich = segs_to_rich(r["expl_segs"])
        d_norm = classify_entry(r["norm_segs"]) if r["norm_segs"] else "unverändert"
        d_expl = classify_entry(r["expl_segs"]) if r["expl_segs"] else "unverändert"
        g = r.get("G", "")
        uncertain = "Ja" if r.get("uncertain") else "Nein"
        auto_notes = list(r.get("notes", []))
        norm_sum = diff_summary(r["norm_segs"])
        if norm_sum:
            auto_notes.append(f"Normtext: {norm_sum}")
        expl_sum = diff_summary(r["expl_segs"])
        if expl_sum:
            auto_notes.append(f"Erläuterung: {expl_sum}")
        notes = "; ".join(auto_notes)

        old_ref = r.get("old_ref", "")
        if r["kind"] == "section_header":
            heading_note = notes or "Überschrift"
            ws.append([r["label"], old_ref, norm_rich, "",
                       d_norm, "", "", "", heading_note])
            row_idx = ws.max_row
            fill = heading_fill if d_norm != "unverändert" else fills["unverändert"]
        else:
            ws.append([r["label"], old_ref, norm_rich, expl_rich,
                       d_norm, d_expl, g, uncertain, notes])
            row_idx = ws.max_row
            worst = max(severity.get(d_norm, 0), severity.get(d_expl, 0))
            if worst == 0:
                d_row = "unverändert"
            elif worst == 3:
                d_row = "geändert"
            else:
                # pick the first non-unverändert
                d_row = d_norm if severity.get(d_norm, 0) >= severity.get(d_expl, 0) else d_expl
            fill = fills.get(d_row, fills["unverändert"])

        for col_i in range(1, 10):
            cell = ws.cell(row=row_idx, column=col_i)
            cell.fill = fill
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    widths = {"A": 20, "B": 20, "C": 70, "D": 70, "E": 18, "F": 18,
              "G": 18, "H": 10, "I": 38}
    for letter, w in widths.items():
        ws.column_dimensions[letter].width = w
    ws.freeze_panes = "A2"

    # Legende
    leg = wb.create_sheet("Legende")
    leg.append(["Spalte / Kennzeichnung", "Bedeutung"])
    leg["A1"].font = header_font
    leg["B1"].font = header_font
    for a, b in [
        ("A – Textziffer",            "neuer Abschnitt + Tz-Nummer, z.B. 'AT 1 Tz. 3'"),
        ("B – alte Referenz",         "alter Pfad der Tz vor Umbenennung/Umnummerierung; leer wenn unverändert. Beispiel: 'AT 4.4.2 Tz. 6' für die heutige 'AT 4.4.2 Tz. 5'"),
        ("C – Normtext",              "Rich-Text des gesamten linken Spaltentexts der Tz"),
        ("D – Erläuterung",           "Rich-Text aller rechten Spalten-Absätze, die zu der Tz gehören"),
        ("E – Änderungsart Normtext",      "unverändert / geändert / gestrichen / hinzugefügt / verschoben"),
        ("F – Änderungsart Erläuterung",   "dasselbe Schema für die Erläuterung"),
        ("G – Verschiebung",          "Heuristische Ziel-/Herkunfts-Tz für inhaltliche Verschiebungen"),
        ("H – Unsicher",              "'Ja' = Verschiebungs-Match unter 75 % Ähnlichkeit"),
        ("I – Anmerkungen",           "Diff-Summary (Wortzahlen, Umformulierungs-Hinweis) und Verschiebungs-Vermerke"),
        ("", ""),
        ("schwarzer Text in C/D",                   "unverändert"),
        ("rot + Durchstreichung",                    "gestrichen (alt)"),
        ("rot + Unterstreichung",                    "neu hinzugefügt"),
        ("grün + Durchstreichung",                   "verschoben (alte Stelle)"),
        ("grün + Unterstreichung",                   "verschoben (neue Stelle)"),
        ("Zeilenhintergrund weiß",                   "Eintrag unverändert"),
        ("Zeilenhintergrund hellblau (#DEEBF7)",     "Tz insgesamt geändert"),
        ("Zeilenhintergrund hellrot (#FCE4E4)",      "Tz gestrichen"),
        ("Zeilenhintergrund hellgrün (#E2EFDA)",     "Tz neu hinzugefügt"),
        ("Zeilenhintergrund hellgelb (#FFF2CC)",     "Verschoben vorher"),
        ("Zeilenhintergrund gelbgrün (#EDFADE)",     "Verschoben nachher"),
        ("Zeilenhintergrund helllila (#EDE7F6)",     "Überschrift geändert"),
    ]:
        leg.append([a, b])
    leg.column_dimensions["A"].width = 44
    leg.column_dimensions["B"].width = 60
    for row in leg.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")

    # Umbenennungen sheet — quick lookup alt → neu.
    if renames:
        un = wb.create_sheet("Umbenennungen")
        un.append(["alter Code", "neuer Code"])
        un["A1"].font = header_font
        un["B1"].font = header_font
        for old, new in renames:
            un.append([old, new])
        un.column_dimensions["A"].width = 18
        un.column_dimensions["B"].width = 18
        un.freeze_panes = "A2"
        for row in un.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(vertical="top")

    wb.save(OUT)


def main():
    doc = fitz.open(PDF)
    all_segs = []
    print(f"Parsing {len(doc)} pages…")
    for pno in range(len(doc)):
        page = doc[pno]
        segs = load_page(page)
        for s in segs:
            s["page"] = pno
        all_segs.extend(segs)
        if (pno + 1) % 20 == 0:
            print(f"  … page {pno+1}/{len(doc)}")

    paragraphs = group_paragraphs(all_segs)

    # Drop everything before the first real section heading.
    first_section = next((k for k, p in enumerate(paragraphs)
                          if p["kind"] == "section"
                          and SECTION_RE.match(plain_text(p["segs"]).strip() or "")),
                         0)
    paragraphs = paragraphs[first_section:]

    # Reading-order per page (y, then col) so Erläuterungen land at the
    # correct Tz.
    by_page = {}
    for p in paragraphs:
        by_page.setdefault(p["page"], []).append(p)
    ordered = []
    for pno in sorted(by_page):
        ordered.extend(sorted(by_page[pno],
                              key=lambda p: (round(p["y_top"], 0),
                                             0 if p["col"] == "L" else 1)))
    paragraphs = ordered

    rows, renames = build_tz_rows(paragraphs)
    print(f"Got {len(rows)} rows "
          f"({sum(1 for r in rows if r['kind']=='section_header')} Überschriften, "
          f"{sum(1 for r in rows if r['kind']=='tz')} Textziffern, "
          f"{len(renames)} Umbenennungen)")

    find_tz_moves(rows)

    write_excel(rows, renames)
    _postprocess_xlsx(OUT)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
