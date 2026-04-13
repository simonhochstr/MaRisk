"""
marisk_parser – PDF-Parser-Bibliothek für die MaRisk-Vergleichsversion.

Stellt die Basis-Funktionen zum Einlesen und Strukturieren der
PDF-Inhalte bereit (Zeichen-Extraktion, Farb-/Strike-/Underline-
Klassifikation, Absatz- und Tz-Gruppierung, Rich-Text-Rendering).

Wird von analyze.py importiert, das daraus die Excel-Datei baut.
"""
import fitz
import re
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

# --- colour classification ---------------------------------------------------
def col_class(c):
    """Map a PyMuPDF int color to: black/red/green/other.

    Purple (~#5C2E91) is used in the Tz number marginal area to mark
    Tz-renumberings (old struck, new underlined). Treated as 'red' so
    the same strike/underline machinery kicks in.
    """
    r = (c >> 16) & 0xFF
    g = (c >> 8) & 0xFF
    b = c & 0xFF
    if r < 60 and g < 60 and b < 60:
        return "black"
    if r > 150 and g < 90 and b < 90:
        return "red"
    if r < 90 and g > 80 and b < 90:
        return "green"
    # Purple (~0x5C2E91 = RGB 92,46,145): Tz renumbering marker.
    if 70 < r < 140 and 20 < g < 80 and 120 < b < 200:
        return "red"
    return "other"

def rect_col_class(rgb):
    if rgb is None: return None
    r, g, b = rgb
    if r > 0.55 and g < 0.4 and b < 0.4:
        return "red"
    if r < 0.35 and g > 0.3 and b < 0.4:
        return "green"
    # Purple fill (~0.36, 0.18, 0.57) for Tz-renumber strike/underline.
    if 0.25 < r < 0.5 and 0.05 < g < 0.3 and 0.45 < b < 0.75:
        return "red"
    return None

# --- formatting classification per char --------------------------------------
# fmt values: "unchanged", "deleted", "added", "moved_from", "moved_to"
def classify_char_run(base_col, has_strike, has_underline):
    if base_col == "black":
        return "unchanged"
    if base_col == "red":
        if has_strike and not has_underline:
            return "deleted"
        if has_underline and not has_strike:
            return "added"
        # ambiguous — assume changed text without mark = body colour only
        return "added" if has_underline else ("deleted" if has_strike else "added")
    if base_col == "green":
        if has_strike and not has_underline:
            return "moved_from"
        if has_underline and not has_strike:
            return "moved_to"
        return "moved_to" if has_underline else "moved_from"
    return "unchanged"

# --- PDF page scan -----------------------------------------------------------
PAGE_MID_X = 360.0         # left-column vs right-column divider
TZ_NUM_X_MAX = 62.0        # left-margin x where Tz numbers appear
TZ_NUM_X_MAX_R = 380.0     # right-column Tz numbers (rare; mostly none)
BODY_LEFT_MIN = 63.0       # left column body starts around here
BODY_RIGHT_MIN = 360.0

_BULLET_RE = re.compile(r"^\s*(?:[-•▪*·]|\(?[a-zA-Z]\)|\(?\d+[.)])(?:\s|$)")

def _starts_bullet(line):
    """True if a peeled line looks like the start of a bullet/enumeration
    item – by the bullet-font of its first segment or the text pattern."""
    sorted_line = sorted(line, key=lambda s: s["x0"])
    head = sorted_line[0]
    if "Courier" in head["font"]:
        # PDF renders bullet dashes in CourierNewPSMT.
        return True
    txt = "".join(s["text"] for s in sorted_line).lstrip()
    return bool(_BULLET_RE.match(txt))


def _sanitise(t):
    """Strip chars Excel can't serialise in inline strings.

    - Private Use Area (U+E000..U+F8FF): custom font glyphs, replace with "-".
    - Forbidden XML control chars (<0x20 except \t\n\r): drop.
    """
    out = []
    for ch in t:
        o = ord(ch)
        if 0xE000 <= o <= 0xF8FF:
            out.append("-")
        elif o < 0x20 and ch not in "\t\n\r":
            continue
        else:
            out.append(ch)
    return "".join(out)

def load_page(page):
    """Return (segments, rects) for a page.

    segments: list of dicts {x0,y0,x1,y1,text,fmt,col,size,font,flags,bold}
              already split so each segment has uniform fmt.
    rects: list of (x0,y0,x1,y1,col) for red/green thin rects.
    """
    rects = []
    for d in page.get_drawings():
        col = rect_col_class(d.get("fill"))
        if col is None:
            col = rect_col_class(d.get("color"))
        if col is None:
            continue
        for it in d.get("items", []):
            if it[0] == "re":
                r = it[1]
                if r.height < 1.6 and r.width > 2 and 80 < r.y0 < 550:
                    rects.append((r.x0, r.y0, r.x1, r.y1, col))
            elif it[0] == "l":
                p1, p2 = it[1], it[2]
                if abs(p1.y - p2.y) < 0.5 and abs(p2.x - p1.x) > 2 and 80 < p1.y < 550:
                    x0, x1 = sorted([p1.x, p2.x])
                    rects.append((x0, p1.y - 0.2, x1, p1.y + 0.2, col))

    rd = page.get_text("rawdict")
    segs = []
    for b in rd["blocks"]:
        if b.get("type") != 0:
            continue
        for line in b["lines"]:
            for s in line["spans"]:
                if not s["chars"]:
                    continue
                txt = "".join(c["c"] for c in s["chars"])
                if not txt.strip() and "chars" in s and len(txt) == 0:
                    continue
                base = col_class(s["color"])
                sbbox = s["bbox"]
                h = sbbox[3] - sbbox[1]
                mid_y = sbbox[1] + 0.55 * h
                bot_y = sbbox[1] + 0.92 * h
                # Only consider rects that overlap this span's y-range and same colour
                rel_rects = [r for r in rects
                             if r[4] == base
                             and r[0] < sbbox[2] and r[2] > sbbox[0]
                             and sbbox[1] - 0.3 < (r[1] + r[3]) / 2 < sbbox[3] + 1.0]

                def char_fmt(cbbox):
                    if base == "black":
                        return "unchanged"
                    cx0, cy0, cx1, cy1 = cbbox
                    cxc = (cx0 + cx1) / 2
                    strike = under = False
                    for rx0, ry0, rx1, ry1, rcol in rel_rects:
                        if rcol != base:
                            continue
                        # Require the char's *center* to sit inside the
                        # rect horizontally. Edge-touching rects (where
                        # rx0 ≈ neighbouring char's x1) must not bleed.
                        if not (rx0 <= cxc <= rx1):
                            continue
                        rmid = (ry0 + ry1) / 2
                        # closer to bottom → underline, closer to mid → strike
                        if abs(rmid - bot_y) < abs(rmid - mid_y) and rmid > mid_y - 0.3:
                            under = True
                        else:
                            strike = True
                    return classify_char_run(base, strike, under)

                # Walk chars, merge runs with same fmt
                run_text = ""
                run_fmt = None
                run_x0 = None
                run_x1 = None
                for ch in s["chars"]:
                    f = char_fmt(ch["bbox"])
                    if run_fmt is None:
                        run_fmt = f
                        run_x0 = ch["bbox"][0]
                    if f != run_fmt:
                        segs.append({
                            "x0": run_x0, "y0": sbbox[1],
                            "x1": run_x1 if run_x1 is not None else ch["bbox"][0],
                            "y1": sbbox[3],
                            "text": run_text, "fmt": run_fmt,
                            "size": s["size"], "font": s["font"],
                            "flags": s["flags"],
                        })
                        run_text = ""
                        run_fmt = f
                        run_x0 = ch["bbox"][0]
                    run_text += ch["c"]
                    run_x1 = ch["bbox"][2]
                if run_text:
                    segs.append({
                        "x0": run_x0, "y0": sbbox[1],
                        "x1": run_x1, "y1": sbbox[3],
                        "text": run_text, "fmt": run_fmt,
                        "size": s["size"], "font": s["font"],
                        "flags": s["flags"],
                    })
    return segs

# --- grouping: lines → paragraphs -------------------------------------------
def group_paragraphs(all_segs):
    """all_segs: list of segs across pages, each with extra 'page' key.

    Returns list of paragraphs:
      {col: 'L'|'R', segs: [...], x_left, y_top, page, kind}
    where kind is one of 'section', 'subsection', 'subheading',
    'tz_body', 'body', 'number'.
    """
    # Filter out header/footer segments (page numbers, bafin info).
    clean = []
    for s in all_segs:
        y = s["y0"]
        if y < 110 or y > 485:
            continue
        if s["size"] >= 7.9 and s["size"] <= 8.1 and "Seite" in s["text"]:
            continue
        clean.append(s)

    # Assign column
    for s in clean:
        xc = (s["x0"] + s["x1"]) / 2
        s["col"] = "L" if xc < PAGE_MID_X else "R"

    # Group into lines on the same page/column by y proximity
    by_page = {}
    for s in clean:
        by_page.setdefault(s["page"], []).append(s)

    paragraphs = []
    for page in sorted(by_page):
        psegs = sorted(by_page[page], key=lambda s: (s["col"], round(s["y0"], 0), s["x0"]))
        # bucket to lines
        lines = []
        cur = []
        cur_col = None
        cur_y = None
        for s in psegs:
            if cur and (s["col"] != cur_col or abs(s["y0"] - cur_y) > 2.5):
                lines.append(cur)
                cur = []
            if not cur:
                cur_col = s["col"]
                cur_y = s["y0"]
            cur.append(s)
        if cur:
            lines.append(cur)

        # line → paragraph: merge consecutive lines in same column if
        # vertical gap is small and the next line's leftmost x aligns
        # with current paragraph's body x (allow hanging indent for Tz numbers).
        # Peel off leading Tz number segments: when a line starts with
        # a segment at x<62 containing only digits, extract it as its
        # own "number" line even if it shares a y-row with body text.
        peeled = []
        for line in lines:
            line_sorted = sorted(line, key=lambda s: s["x0"])
            head = line_sorted[0]
            if (head["col"] == "L"
                    and head["x0"] < TZ_NUM_X_MAX
                    and re.fullmatch(r"\d+[a-z]?", head["text"].strip())):
                # Peel off ALL adjacent digit segments at the left
                # margin (a renumbered Tz renders as two segments like
                # "5"=deleted + "4"=added, both still inside the number
                # column; only the first one would otherwise land in the
                # peeled number line).
                k = 1
                while (k < len(line_sorted)
                       and line_sorted[k]["x0"] < TZ_NUM_X_MAX + 6
                       and re.fullmatch(r"\d+[a-z]?",
                                        line_sorted[k]["text"].strip())):
                    k += 1
                peeled.append(line_sorted[:k])
                rest = line_sorted[k:]
                if rest:
                    peeled.append(rest)
            else:
                peeled.append(line_sorted)
        lines = peeled

        i = 0
        while i < len(lines):
            line = lines[i]
            col = line[0]["col"]
            xs = [s["x0"] for s in line]
            min_x = min(xs)
            max_y = max(s["y1"] for s in line)
            # classify line
            line_text = "".join(s["text"] for s in line).strip()
            big = any(s["size"] >= 11.5 for s in line)
            med = any(11.0 <= s["size"] < 11.5 for s in line)
            bold_only = all((s["flags"] & 16) for s in line) and not big and not med
            is_num = (col == "L"
                      and min_x < TZ_NUM_X_MAX
                      and re.fullmatch(r"\d+[a-z]?", line_text) is not None)
            kind = None
            if big:
                kind = "section"
            elif med:
                kind = "subsection"
            elif is_num:
                kind = "number"
            elif (col == "L"
                    and bold_only
                    and re.match(r"^(AT|BT|BTO|BTR)\s?\d+(?:\.\d+)+",
                                 line_text, re.IGNORECASE)):
                # Deeper subsection headings (e.g. "AT 4.4.1") are set
                # in SegoeUI-Bold ~9–10 pt instead of Cambria, so they
                # slip past the size-based filters. Promote them.
                kind = "subsection"
            else:
                # Bold one-liners inside a column are *not* promoted to a
                # separate paragraph kind anymore – they belong to the
                # surrounding Textziffer / Erläuterung and must merge with
                # the following body lines. The bold flag is still carried
                # on the spans, so the rich-text output keeps them bold.
                kind = "body"

            para_segs = list(line)
            # merge following body lines into same paragraph
            if kind == "body":
                j = i + 1
                while j < len(lines):
                    nxt = lines[j]
                    if nxt[0]["col"] != col:
                        break
                    nxt_min_x = min(s["x0"] for s in nxt)
                    nxt_y = min(s["y0"] for s in nxt)
                    nxt_text = "".join(s["text"] for s in nxt).strip()
                    if any(s["size"] >= 11 for s in nxt):
                        break
                    if (col == "L" and nxt_min_x < TZ_NUM_X_MAX
                            and re.fullmatch(r"\d+[a-z]?", nxt_text) is not None):
                        break
                    if nxt_y - max_y > 10:
                        break
                    # If this line starts a bullet / enumeration item,
                    # insert a line break marker so the rich-text cell
                    # shows the list on separate lines.
                    if _starts_bullet(nxt):
                        para_segs.append({"text": "\n", "fmt": "unchanged",
                                          "flags": 0, "size": line[0]["size"],
                                          "font": line[0]["font"],
                                          "x0": nxt_min_x, "y0": nxt_y,
                                          "x1": nxt_min_x, "y1": nxt_y})
                    para_segs.extend(nxt)
                    max_y = max(s["y1"] for s in nxt)
                    j += 1
                paragraphs.append({
                    "col": col, "segs": para_segs, "kind": kind,
                    "x_left": min_x, "y_top": line[0]["y0"], "page": page,
                })
                i = j
                continue

            paragraphs.append({
                "col": col, "segs": para_segs, "kind": kind,
                "x_left": min_x, "y_top": line[0]["y0"], "page": page,
            })
            i += 1

    # ---- cross-page merging ------------------------------------------------
    # If a page's first body paragraph on the left column starts very high
    # (near the top of the content area) and the previous page's last
    # paragraph in the same column was also a body paragraph near the bottom,
    # merge them so a Tz that continues across a page break ends up as one
    # paragraph.
    merged = []
    for p in paragraphs:
        if (merged
                and p["kind"] == "body"
                and merged[-1]["kind"] == "body"
                and merged[-1]["col"] == p["col"]
                and p["page"] == merged[-1]["page"] + 1
                and p["y_top"] < 140
                and max(s["y1"] for s in merged[-1]["segs"]) > 430):
            merged[-1]["segs"].extend(p["segs"])
            continue
        merged.append(p)
    return merged

# --- rich-text + entry classification ---------------------------------------
FMT_COLOR = {
    "unchanged":  "FF000000",
    "deleted":    "FFC00000",
    "added":      "FFC00000",
    "moved_from": "FF2C6234",
    "moved_to":   "FF2C6234",
}

def segs_to_rich(segs):
    """Build an openpyxl CellRichText from segments."""
    parts = []
    for s in segs:
        t = _sanitise(s["text"])
        if not t:
            continue
        fmt = s["fmt"]
        color = FMT_COLOR.get(fmt, "000000")
        strike = fmt in ("deleted", "moved_from")
        under  = "single" if fmt in ("added", "moved_to") else "none"
        bold   = bool(s.get("flags", 0) & 16)
        inline = InlineFont(
            rFont="Calibri", sz="10", color=color,
            b=bold, strike=strike, u=under,
        )
        parts.append(TextBlock(inline, t))
    if not parts:
        return CellRichText("")
    return CellRichText(*parts)

def classify_entry(segs):
    fmts = {s["fmt"] for s in segs if s["text"].strip()}
    if not fmts or fmts == {"unchanged"}:
        return "unverändert"
    if fmts == {"deleted"}:
        return "gestrichen"
    if fmts == {"added"}:
        return "hinzugefügt"
    if fmts == {"moved_from"}:
        return "verschoben vorher"
    if fmts == {"moved_to"}:
        return "verschoben nachher"
    return "geändert"

def plain_text(segs, fmt_filter=None):
    return "".join(s["text"] for s in segs
                   if fmt_filter is None or s["fmt"] in fmt_filter)


def diff_summary(segs):
    """Generate a human-readable change summary for Spalte G.

    Counts inserted / deleted / moved words and prefixes them with
    short German sentences. Returns '' for entirely unchanged content.
    """
    def words(fmts):
        return len(plain_text(segs, fmt_filter=fmts).split())
    n_del = words({"deleted"})
    n_add = words({"added"})
    n_mv_from = words({"moved_from"})
    n_mv_to = words({"moved_to"})
    n_unchanged = words({"unchanged"})
    bits = []
    if n_del == 0 and n_add == 0 and n_mv_from == 0 and n_mv_to == 0:
        return ""
    if n_del and n_add:
        bits.append(f"{n_del} Wort(e) gestrichen, {n_add} neu")
    elif n_del:
        bits.append(f"{n_del} Wort(e) gestrichen")
    elif n_add:
        bits.append(f"{n_add} Wort(e) neu")
    if n_mv_from or n_mv_to:
        bits.append(f"{n_mv_from + n_mv_to} Wort(e) verschoben")
    # Rough rewrite heuristic: if both sides are present and roughly
    # the same size, likely a reformulation rather than net change.
    if n_del and n_add and abs(n_del - n_add) <= max(2, 0.3 * max(n_del, n_add)):
        bits.append("wahrscheinlich umformuliert, Substanz ggf. unverändert")
    return "; ".join(bits)

# --- structural walk: derive Tz-label per paragraph -------------------------
SECTION_RE = re.compile(r"^(AT|BT|BTO|BTR)\s?\d+(?:\.\d+)*", re.IGNORECASE)


def split_footnote_from_tz(text, current_tz):
    """Disentangle a number blob like "34" (= fn 3 + Tz 4) from a real Tz.

    The MaRisk PDF renders a footnote reference marker at the same x
    position and same font size as the Tz number, so two-digit blobs
    like "32", "1110" appear in the left margin. Empirically the
    footnote marker sits *before* the Tz number in the span.

    Returns (tz_number_str, footnote_prefix).
      tz_number_str is None if the entire blob is a pure footnote
      marker (no embedded real Tz number, like "75" after Tz 6).
    """
    if not text.isdigit():
        return text, ""

    if current_tz is None or not str(current_tz).isdigit():
        expected = ["1", "2"]           # first Tz of a section
    else:
        n = int(current_tz)
        expected = [str(n + 1), str(n + 2)]   # next, or a small gap

    # Direct match first.
    if text in expected:
        return text, ""
    # Suffix match – the last k chars form the real Tz, prefix is footnote.
    for exp in expected:
        if text.endswith(exp) and len(text) > len(exp):
            return exp, text[:-len(exp)]
    # Neither matches – pure footnote marker.
    return None, text

def _postprocess_xlsx(path):
    """Add xml:space='preserve' to every <t> element in sheet XML parts.

    openpyxl's CellRichText serializer omits this attribute on runs
    whose text is pure whitespace, which makes Excel report a
    "Repaired Records: String properties" warning on open. Rewriting
    the zipped sheet XML fixes it without affecting any cell content.
    """
    import zipfile, shutil, re, os
    tmp = path + ".tmp"
    with zipfile.ZipFile(path, "r") as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                text = data.decode("utf-8")
                # Replace <t>...</t> and <t xml:space="..."> variants
                # with a canonical xml:space="preserve" form.
                text = re.sub(r"<t(?:\s+xml:space=\"[^\"]*\")?>", r'<t xml:space="preserve">', text)
                data = text.encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, path)

