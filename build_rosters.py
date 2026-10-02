#!/usr/bin/env python3
"""
build_rosters.py
================
Generate F.A.S.T. Rescue course rosters (Word .docx) from an Excel export.

Usage:
    python build_rosters.py EXPORT.xlsx [--templates TEMPLATE_DIR] [--out OUTPUT_DIR]

For every data row in the export it:
  1. Detects the course type (In-Class / Blended / Recertification) from the Name column.
  2. Picks the matching blank template *by its layout*, not its filename
     (the In-Class / Blended template files are cross-named, so we identify by content).
  3. Fills the info table and the participant table, preserving all template formatting.
  4. Saves "<Instructor> - <Start date> - <Company>.docx".
"""

import argparse
import copy
import glob
import os
import re
import sys
from datetime import datetime, time as dtime

from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ----------------------------------------------------------------------------- #
# Field sanitization
# ----------------------------------------------------------------------------- #

def _strip_acct(s: str) -> str:
    """Strip leading and trailing account numbers from a string."""
    s = re.sub(r"^\s*\d+\s*", "", s)
    s = re.sub(r"\s*\d+\s*$", "", s)
    return re.sub(r"\s+", " ", s).strip(" :")


def clean_company(customer: str, name_col: str = "") -> str:
    """Extract the parent customer name.  Prefers the first colon-segment of the
    Name column (Col A) which always has the parent customer, falling back to
    the Customer column (Col B) when Col A is absent or unhelpful.

    Col A pattern: '<code> Parent Customer : [<subcode> Sub Customer :] <date> Course...'
    Col B pattern: '<code>[:<subcode>] Customer Name'
    """
    # Try Col A first — the first segment before a ' : ' is always the parent
    a = (name_col or "").strip()
    if " : " in a:
        first_seg = a.split(" : ")[0].strip()
        parent = _strip_acct(first_seg)
        if parent:
            return parent

    # Fallback to Col B
    c = (customer or "").strip()
    if ":" in c:
        segs = [s.strip() for s in c.split(":")]
        c = max(segs, key=lambda s: sum(ch.isalpha() for ch in s))
    return _strip_acct(c)


def _strip_trailing_city(text: str, location: str) -> str:
    """Remove trailing city words (e.g. 'North York') by matching against the
    words found in the Course Location."""
    loc_words = {w.lower() for w in re.findall(r"[A-Za-z]{2,}",
                 (location or "").replace("_x000D_", " "))}
    if not loc_words:
        return text
    words = text.split(" ")
    while len(words) > 1 and words[-1].strip(".,").lower() in loc_words:
        words.pop()
    return " ".join(words)


def extract_course_name(name_col: str, location: str = "") -> str:
    """
    Clean course name from the Name column.  The 8-digit YYYYMMDD date in the
    string is the anchor — everything after it is the course name (+ trailing
    city).  Always ends with 'Training'.
    """
    raw = (name_col or "").strip()
    # find the 8-digit date — works regardless of how many colons precede it
    m = re.search(r"\b(\d{8})\s+", raw)
    if m:
        after = raw[m.end():]
    else:
        # fallback: try the old colon split
        after = raw.split(" : ", 1)[1] if " : " in raw else raw
        after = re.sub(r"^\s*\d{8}\s+", "", after).strip()
    # keep only the first line (drops _x000D_ multi-line annotations)
    after = re.split(r"_x000D_|[\r\n]", after)[0].strip()
    after = re.sub(r"\s+", " ", after)

    m2 = re.search(r"\bTraining\b", after, flags=re.I)
    if m2:                                          # 'Training' marks the end
        core = after[:m2.start()].strip()
    else:                                          # no 'Training' -> drop city
        core = _strip_trailing_city(after, location).strip()
    core = re.sub(r"\bRecertificatio(n)?\b", "Recertification", core, flags=re.I)
    core = core.rstrip(" -,&").strip()
    if not core:
        core = after
    if not re.search(r"\bTraining$", core, flags=re.I):
        core = f"{core} Training"
    return core


def extract_location(location: str, company: str, contact: str = "") -> str:
    """Strip a leading contact name and/or company name off the address."""
    loc = (location or "").strip()
    prefixes = [p for p in (contact, company) if p]
    again = True
    while again:
        again = False
        for pre in prefixes:
            if loc.lower().startswith(pre.lower()):
                loc = loc[len(pre):].lstrip(" ,-")
                again = True
    return re.sub(r"\s+", " ", loc).strip()


def extract_contact_name(contact_col: str) -> str:
    """'519817 CUPE Ontario : Olivia Kirby' -> 'Olivia Kirby'."""
    raw = (contact_col or "").strip()
    return raw.rsplit(" : ", 1)[-1].strip() if " : " in raw else raw


def format_contact(name: str, phone) -> str:
    """'Name – phone' (en-dash). Phone passed through verbatim."""
    phone = "" if phone is None else str(phone).strip()
    name = (name or "").strip()
    if name and phone:
        return f"{name} \u2013 {phone}"
    return name or phone


def _as_dt(v):
    if isinstance(v, datetime):
        return v
    if isinstance(v, str) and v.strip():
        for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%Y%m%d"):
            try:
                return datetime.strptime(v.strip(), fmt)
            except ValueError:
                pass
    return None


def _fmt_single(d) -> str:
    return f"{d:%B} {d.day}, {d.year}" if d else ""


def _first_date(start, other, end):
    """Earliest valid session date (drops dates before the start, e.g. typos)."""
    s = _as_dt(start)
    collected = [d for d in (_as_dt(start), _as_dt(other), _as_dt(end)) if d]
    if not collected:
        return None
    anchor = s or min(collected)
    valid = [d for d in collected if d.date() >= anchor.date()]
    return min(valid) if valid else anchor


def format_dates(start, other, end) -> str:
    """
    List the actual session dates from Course Date / Other Date / Course End Date.
      single day       -> 'May 2, 2026'
      two days         -> 'May 28 & 29, 2026'
      several sessions -> 'May 2, 3 & 5, 2026'   (month repeated only when it changes)
    Dates earlier than the start (e.g. a typo'd year) are dropped.
    """
    s = _as_dt(start)
    collected = [d for d in (_as_dt(start), _as_dt(other), _as_dt(end)) if d]
    if not collected:
        return ""
    anchor = s or min(collected)
    uniq = sorted({d.date() for d in collected if d.date() >= anchor.date()})
    if len(uniq) == 1:
        d = uniq[0]
        return f"{d.strftime('%B')} {d.day}, {d.year}"

    multi_year = len({d.year for d in uniq}) > 1
    pieces, prev_month = [], None
    for d in uniq:
        if multi_year:
            pieces.append(f"{d.strftime('%B')} {d.day}, {d.year}")
        elif d.month != prev_month:
            pieces.append(f"{d.strftime('%B')} {d.day}")
        else:
            pieces.append(str(d.day))
        prev_month = d.month
    body = pieces[0] if len(pieces) == 1 else ", ".join(pieces[:-1]) + " & " + pieces[-1]
    return body if multi_year else f"{body}, {uniq[0].year}"


def _as_time(v):
    if isinstance(v, dtime):
        return v
    if isinstance(v, datetime):
        return v.time()
    if isinstance(v, str) and v.strip():
        for fmt in ("%H:%M", "%I:%M %p", "%I:%M%p"):
            try:
                return datetime.strptime(v.strip(), fmt).time()
            except ValueError:
                pass
    return None


def _fmt_t(t) -> str:
    # 8:30 am  (no leading zero, lowercase meridiem)
    return f"{(t.hour % 12) or 12}:{t.minute:02d} {'am' if t.hour < 12 else 'pm'}"


def format_time(t_from, t_to) -> str:
    a, b = _as_time(t_from), _as_time(t_to)
    if a and b:
        return f"{_fmt_t(a)} to {_fmt_t(b)}"
    return _fmt_t(a) if a else ""


BULLET_RE = re.compile(r"^\s*[\u2022\u00b7\u25aa\u2023\u25e6\*\-]\s+")
EMAILISH = re.compile(r"[^\s@]+@[^\s@]+")


def _is_email_tok(tok: str) -> bool:
    return bool(EMAILISH.search(tok))


def _is_header(line: str) -> bool:
    """Group/section header lines that are not participants."""
    l = line.strip()
    if not l:
        return False
    if re.search(r"group\s*\(", l, re.I):
        return True
    if l.endswith(":"):
        return True
    return False


def _invert_comma(name: str) -> str:
    """'Hwang, Mikaela' -> 'Mikaela Hwang'."""
    if name.count(",") == 1:
        last, first = [p.strip() for p in name.split(",")]
        if last and first:
            return f"{first} {last}"
    return name


def _finalize(raw_name: str):
    """Collapse spaces, split a trailing '- note', invert 'Last, First'."""
    parts = re.split(r"-\s+", raw_name, maxsplit=1)
    name = re.sub(r"\s+", " ", parts[0]).strip(" \t-\u2022\u00b7")
    name = re.sub(r"^\d+\s*[\.\):]?\s+", "", name)   # drop any leading "1." / "1)" / "1 "
    note = re.sub(r"\s+", " ", parts[1]).strip() if len(parts) > 1 else ""
    name = _invert_comma(name)
    return name, note


def _names_from_block(text: str):
    """
    Auto-detect the participant format of one text block and return raw name strings.
      * has emails   -> accumulate lines until an email terminates a person
      * else bullets -> each bullet marks a new person
      * else         -> one person per line
    Leading bullets, emails, blank lines and group headers are handled in every mode.
    """
    lines = [l.strip() for l in re.split(r"[\r\n]+", text.replace("_x000D_", ""))]
    has_email = any("@" in l for l in lines)
    has_bullet = any(BULLET_RE.match(l) for l in lines)
    people, cur = [], []

    def flush():
        if cur:
            people.append(" ".join(cur))
            cur.clear()

    if has_email:
        for l in lines:
            if not l:
                flush(); continue
            if _is_header(l):
                flush(); continue
            l = BULLET_RE.sub("", l)
            toks = l.split()
            cur.extend(t for t in toks if not _is_email_tok(t))
            if any(_is_email_tok(t) for t in toks):
                flush()
        flush()
    elif has_bullet:
        for l in lines:
            if not l:
                flush(); continue
            if _is_header(l):
                flush(); continue
            if BULLET_RE.match(l):
                flush()
                l = BULLET_RE.sub("", l)
            cur.extend(t for t in l.split() if not _is_email_tok(t))
        flush()
    else:
        for l in lines:
            if not l or _is_header(l):
                continue
            people.append(BULLET_RE.sub("", l))
    return people


def parse_participants(part_col, recert_col, sort_alpha: bool = True):
    """Returns a list of (name, note) tuples, format auto-detected per block."""
    rows = []
    for blob in (part_col, recert_col):
        if not blob:
            continue
        for raw in _names_from_block(str(blob)):
            name, note = _finalize(raw)
            if name:
                rows.append((name, note))
    if sort_alpha:
        rows.sort(key=lambda r: r[0].lower())
    return rows


def parse_participants_tagged(part_col, recert_col, sort_alpha: bool = True):
    """Like parse_participants but returns (name, note, source) where source is
    'participant' or 'recert'."""
    rows = []
    for blob, tag in ((part_col, "participant"), (recert_col, "recert")):
        if not blob:
            continue
        for raw in _names_from_block(str(blob)):
            name, note = _finalize(raw)
            if name:
                rows.append((name, note, tag))
    if sort_alpha:
        rows.sort(key=lambda r: r[0].lower())
    return rows


# ----------------------------------------------------------------------------- #
# Template handling
# ----------------------------------------------------------------------------- #

def detect_course_type(name_col: str) -> str:
    low = (name_col or "").lower()
    if "lift truck" in low or "forklift" in low:
        return "lift_truck"
    if "elevated work platform" in low or "elevating work platform" in low or "ewp" in low:
        return "ewp"
    if "working at height" in low:
        return "working_at_heights"
    if "recert" in low:
        if "blended" in low:
            return "blended"         # blended + recert combo -> blended template
        if "in-class" in low or "in class" in low:
            return "in-class"        # in-class + recert combo -> in-class template
        return "recert"
    if "blended" in low:
        return "blended"
    if "in-class" in low or "in class" in low:
        return "in-class"
    return "general"


def _course_flags(name_col: str) -> dict:
    """Detect sub-type flags for special participant routing."""
    low = (name_col or "").lower()
    is_wah = "working at height" in low
    is_recert = "recert" in low or "refresher" in low
    is_blended = "blended" in low
    is_inclass = "in-class" in low or "in class" in low
    is_full = "full" in low
    return {
        "wah": is_wah,
        "wah_recert_only": is_wah and is_recert and not is_full,
        "wah_full_only": is_wah and not is_recert,
        "wah_mix": is_wah and is_recert and is_full,
        "blended_recert": is_blended and is_recert,
        "inclass_recert": is_inclass and is_recert,
    }


def classify_template(path: str) -> str:
    """Identify a template. New templates are matched by filename; the First Aid
    trio is matched by content (their In-Class/Blended filenames are unreliable)."""
    base = re.sub(r"[_\-]+", " ", os.path.basename(path).lower())
    if "lift truck" in base or "forklift" in base:
        return "lift_truck"
    if "ewp" in base or "elevated work" in base:
        return "ewp"
    if "working at height" in base:
        return "working_at_heights"
    if "general" in base:
        return "general"
    d = Document(path)
    header = [c.text.strip().lower() for c in d.tables[1].rows[0].cells]
    if any("online part" in h for h in header):
        return "blended"
    material = ""
    for row in d.tables[0].rows:
        if row.cells[0].text.strip().lower() == "course material":
            material = row.cells[1].text.lower()
            break
    return "in-class" if "manual" in material else "recert"


def load_template_map(template_dir: str) -> dict:
    mapping = {}
    for path in glob.glob(os.path.join(template_dir, "*.docx")):
        if "template" not in os.path.basename(path).lower():
            continue
        try:
            mapping[classify_template(path)] = path
        except Exception as exc:                       # noqa: BLE001
            print(f"  ! could not classify {os.path.basename(path)}: {exc}",
                  file=sys.stderr)
    return mapping


# ----------------------------------------------------------------------------- #
# Word writing helpers
# ----------------------------------------------------------------------------- #

def _strip_list_p(p_el):
    """Remove auto-list numbering / list indentation from a <w:p> element."""
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return
    for tag in ("w:numPr", "w:ind"):
        el = pPr.find(qn(tag))
        if el is not None:
            pPr.remove(el)
    ps = pPr.find(qn("w:pStyle"))
    if ps is not None and ps.get(qn("w:val"), "").lower().startswith("list"):
        pPr.remove(ps)


def set_cell(cell, text, bold=False, font="Times New Roman", size=12, center=False):
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    _strip_list_p(p._p)
    if center:
        p.alignment = 1  # WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.font.name = font
    run.font.size = Pt(size)
    run.bold = bold
    return run


def strip_table_numbering(table):
    """Remove auto-list numbering from every data cell (keeps header row)."""
    for row in table.rows[1:]:
        for cell in row.cells:
            for p in cell.paragraphs:
                _strip_list_p(p._p)


def _blank_row_like(table, ref_tr):
    new_tr = copy.deepcopy(ref_tr)
    table._tbl.append(new_tr)
    row = table.rows[-1]
    for c in row.cells:
        for p in c.paragraphs:
            for r in list(p.runs):
                r._element.getparent().remove(r._element)
    return row


def _set_tc_width(tc, w):
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:type"), "dxa")
    tcW.set(qn("w:w"), str(int(w)))


def add_number_column(table, header="Sr. No.", width=850):
    """Prepend a 'Sr. No.' column and number every data row 1..N."""
    grid = table._tbl.find(qn("w:tblGrid"))
    gcols = grid.findall(qn("w:gridCol"))
    name_w = int(gcols[0].get(qn("w:w")))
    new_name_w = max(name_w - width, 600)
    gcols[0].set(qn("w:w"), str(new_name_w))
    new_gc = OxmlElement("w:gridCol")
    new_gc.set(qn("w:w"), str(width))
    grid.insert(0, new_gc)

    for row in table.rows:                       # clone each row's own first cell
        first_tc = row._tr.findall(qn("w:tc"))[0]
        new_tc = copy.deepcopy(first_tc)
        paras = new_tc.findall(qn("w:p"))
        for extra in paras[1:]:                  # keep a single, clean paragraph
            new_tc.remove(extra)
        p0 = new_tc.find(qn("w:p"))
        for r in p0.findall(qn("w:r")):
            p0.remove(r)
        _strip_list_p(p0)
        _set_tc_width(new_tc, width)
        _set_tc_width(first_tc, new_name_w)
        first_tc.addprevious(new_tc)

    n = 0
    for ri, row in enumerate(table.rows):        # write header then 1..N
        if ri == 0:
            set_cell(row.cells[0], header, bold=True, center=True)
        else:
            n += 1
            set_cell(row.cells[0], str(n), center=True)


def fill_info_table(doc, info: dict):
    table = doc.tables[0]
    for row in table.rows:
        label = row.cells[0].text.strip()
        if label in info:
            set_cell(row.cells[1], info[label])


def ensure_data_rows(table, target):
    """Make the participant table have exactly `target` data rows (excl. header)."""
    data = table.rows[1:]
    cur = len(data)
    if cur < target:
        ref_tr = data[-1]._tr if data else table.rows[-1]._tr
        for _ in range(target - cur):
            _blank_row_like(table, ref_tr)
    elif cur > target:
        for _ in range(cur - target):
            tr = table.rows[-1]._tr
            tr.getparent().remove(tr)


def _count_header_rows(table):
    """Count header rows — rows where at least one cell is non-empty and looks
    like a label rather than data.  For landscape templates with 3-row headers
    (merged spans), this detects them correctly."""
    # simple heuristic: header rows are the rows before the first all-empty row
    for ri, row in enumerate(table.rows):
        if ri == 0:
            continue  # always a header
        vals = [c.text.strip() for c in row.cells]
        if all(v == "" or v == vals[0] for v in vals) and vals[0] == "":
            return ri
    return 1  # fallback: only row 0 is header


def _has_sr_no(table):
    """Check if the first column is already a Sr. No. column."""
    c0 = table.rows[0].cells[0].text.strip().lower()
    return c0 in ("sr. no.", "sr.no.", "sr no", "#", "no.", "s.no.", "s. no.")


def _fix_row_heights(table, header_rows):
    """Set minimum row height on data rows — constant unless a name wraps."""
    from docx.oxml.ns import qn as _qn
    ref_h = None
    for ri in range(header_rows, len(table.rows)):
        trPr = table.rows[ri]._tr.find(_qn("w:trPr"))
        if trPr is not None:
            trH = trPr.find(_qn("w:trHeight"))
            if trH is not None:
                if ref_h is None:
                    ref_h = trH.get(_qn("w:val"))
                trH.set(_qn("w:hRule"), "atLeast")
                trH.set(_qn("w:val"), ref_h)


def _find_col(header, *keywords):
    """Find column index by keyword in header."""
    for i, h in enumerate(header):
        for kw in keywords:
            if kw in h:
                return i
    return None


def fill_participants(doc, participants, flags=None):
    """Fill participant table.  `participants` is a list of (name, note, source)
    tuples where source is 'participant' or 'recert'.  `flags` controls special
    column routing for WAH, blended+recert, in-class+recert."""
    flags = flags or {}
    table = doc.tables[1]
    has_sr = _has_sr_no(table)
    hdr_count = _count_header_rows(table)
    header = [c.text.strip().lower() for c in table.rows[0].cells]
    name_idx = 1 if has_sr else 0
    note_idx = _find_col(header, "note")
    wahf_idx = _find_col(header, "wahf", "wahr")
    online_idx = _find_col(header, "online part")

    target = max(20, len(participants) + 5)
    data_start = hdr_count
    cur_data = len(table.rows) - data_start
    if cur_data < target:
        ref_tr = table.rows[-1]._tr
        for _ in range(target - cur_data):
            _blank_row_like(table, ref_tr)
    elif cur_data > target:
        for _ in range(cur_data - target):
            tr = table.rows[-1]._tr
            tr.getparent().remove(tr)

    data_rows = list(table.rows[data_start:])
    for i, entry in enumerate(participants):
        name, note = entry[0], entry[1]
        source = entry[2] if len(entry) > 2 else "participant"
        row = data_rows[i]
        set_cell(row.cells[name_idx], name)

        # --- WAH: WAHF/WAHR column ---
        if flags.get("wah") and wahf_idx is not None:
            if flags.get("wah_full_only"):
                set_cell(row.cells[wahf_idx], "WAHF")
            elif flags.get("wah_recert_only"):
                set_cell(row.cells[wahf_idx], "WAHR")
            elif flags.get("wah_mix"):
                set_cell(row.cells[wahf_idx],
                         "WAHR" if source == "recert" else "WAHF")

        # --- Blended + Recert: Online Part 1 = "Recert" for recert ppl ---
        if flags.get("blended_recert") and online_idx is not None:
            if source == "recert":
                set_cell(row.cells[online_idx], "Recert")

        # --- In-Class + Recert: Instructor Notes = "Recert" for recert ppl ---
        if flags.get("inclass_recert") and note_idx is not None:
            if source == "recert":
                note = "Recert" if not note else f"Recert - {note}"

        if note and note_idx is not None:
            set_cell(row.cells[note_idx], note)

    # fill sr. no. numbers
    if has_sr:
        for i, row in enumerate(data_rows):
            set_cell(row.cells[0], str(i + 1), center=True)
    else:
        strip_table_numbering(table)
        add_number_column(table)

    _fix_row_heights(table, hdr_count)


_MINOR_WORDS = {"a", "an", "and", "as", "at", "but", "by", "for", "if", "in",
                "into", "nor", "of", "on", "onto", "or", "the", "to", "via",
                "vs", "with"}


def _title_segment(seg: str) -> str:
    words = seg.split(" ")
    n = len(words)
    out = []
    for i, w in enumerate(words):
        ai = next((k for k, ch in enumerate(w) if ch.isalpha()), None)
        if ai is None:                       # no letters (numbers, "&", etc.)
            out.append(w)
            continue
        is_minor = w[ai:].strip(".,()").lower() in _MINOR_WORDS
        if is_minor and 0 < i < n - 1:       # lowercase minor words mid-segment
            out.append(w[:ai] + w[ai:].lower())
        else:                                # capitalise first letter, keep the rest
            out.append(w[:ai] + w[ai].upper() + w[ai + 1:])
    return " ".join(out)


def _cap_words(s: str) -> str:
    """Title-case each ' - ' segment: first letter of each word capitalised,
    acronyms preserved, minor words (of, the, and, ...) kept lowercase unless
    first or last word of their segment."""
    return " - ".join(_title_segment(seg) for seg in s.split(" - "))


def safe_filename(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "", s).strip()


# ----------------------------------------------------------------------------- #
# Main
# ----------------------------------------------------------------------------- #

H = {  # normalised header -> our key
    "name": "name", "customer": "customer", "course location": "location",
    "training contact": "contact", "training contact phone": "phone",
    "course date": "start", "course end date": "end", "other date": "other",
    "time (from)": "tfrom", "time (to)": "tto", "instructors": "instructor",
    "participant names": "participants",
    "recertification participants": "recert_participants",
    "on-site contact": "onsite_contact", "on-site phone": "onsite_phone",
}


def read_export(xlsx_path):
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    headers = {}
    for c in range(1, ws.max_column + 1):
        h = (ws.cell(row=1, column=c).value or "").strip().lower()
        if h in H:
            headers[H[h]] = c
    rows = []
    for r in range(2, ws.max_row + 1):
        if not any(ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)):
            continue
        rows.append({k: ws.cell(row=r, column=col).value for k, col in headers.items()})
    return rows


_BLANK_INSTRUCTORS = {"noeline", "frank keegan"}


def _should_blank_instructor(name: str) -> bool:
    low = (name or "").lower()
    return any(b in low for b in _BLANK_INSTRUCTORS)


def build_one(rec, template_map, out_dir, sort_alpha=True, include_cancelled=False,
              fname_suffix=""):
    name_col = rec.get("name") or ""
    company = clean_company(rec.get("customer"), name_col)

    if not include_cancelled and "cancel" in name_col.lower():
        return ("skipped", "course marked CANCELLED", None, 0)

    ctype = detect_course_type(name_col)
    tpl = template_map.get(ctype)
    if not tpl:
        return ("skipped", f"no template for '{ctype}'", ctype, 0)

    instructor_raw = (rec.get("instructor") or "").strip()
    instructor_display = "" if _should_blank_instructor(instructor_raw) else instructor_raw

    contact_name = extract_contact_name(rec.get("contact"))
    contact_str = format_contact(contact_name, rec.get("phone"))
    onsite_name = extract_contact_name(rec.get("onsite_contact"))
    onsite_phone = rec.get("onsite_phone")
    if onsite_name:
        onsite_str = format_contact(onsite_name, onsite_phone)
        contact_str = f"{contact_str} / {onsite_str}" if contact_str else onsite_str
    info = {
        "Course Name": extract_course_name(name_col, rec.get("location")),
        "Company Name": company,
        "Location": extract_location(rec.get("location"), company, contact_name),
        "Contact": contact_str,
        "Date": format_dates(rec.get("start"), rec.get("other"), rec.get("end")),
        "Time": format_time(rec.get("tfrom"), rec.get("tto")),
        "Instructor": instructor_display,
    }
    flags = _course_flags(name_col)

    # Decide which columns to read based on course type
    part_col = rec.get("participants")
    recert_col = rec.get("recert_participants")

    if flags["wah"] and flags["wah_full_only"]:
        # WAH full only: use Participant Names only
        participants = parse_participants_tagged(part_col, None, sort_alpha)
    elif flags["wah"] and flags["wah_recert_only"]:
        # WAH recert only: use Recertification Participants only
        participants = parse_participants_tagged(None, recert_col, sort_alpha)
    else:
        # Everything else (including WAH mix, blended+recert, in-class+recert):
        # use both columns
        participants = parse_participants_tagged(part_col, recert_col, sort_alpha)

    doc = Document(tpl)
    fill_info_table(doc, info)
    fill_participants(doc, participants, flags)

    start_label = _fmt_single(_first_date(rec.get("start"), rec.get("other"), rec.get("end")))
    instructor_fname = instructor_raw or "Instructor"
    parts = [p for p in (instructor_fname, start_label, company) if p]
    if fname_suffix:
        parts.append(fname_suffix)
    raw = " - ".join(parts)
    fname = safe_filename(_cap_words(raw).rstrip(" .")) + ".docx"
    out_path = os.path.join(out_dir, fname)
    doc.save(out_path)
    return ("ok", out_path, ctype, len(participants))


def _assign_suffixes(rows):
    """For same customer + same date, determine filename suffixes:
       - different time   -> 'Session 1', 'Session 2', …
       - different course -> course name appended
       - different location -> location appended
       Returns a list of suffix strings, one per row."""
    from collections import defaultdict

    def _date_key(rec):
        d = _as_dt(rec.get("start"))
        return d.date().isoformat() if d else ""

    def _time_key(rec):
        return format_time(rec.get("tfrom"), rec.get("tto"))

    def _course_key(rec):
        return extract_course_name(rec.get("name") or "", rec.get("location") or "")

    def _loc_key(rec):
        company = clean_company(rec.get("customer"), rec.get("name"))
        contact = extract_contact_name(rec.get("contact"))
        return extract_location(rec.get("location"), company, contact)

    # group by (company, date)
    groups = defaultdict(list)
    for i, rec in enumerate(rows):
        company = clean_company(rec.get("customer"), rec.get("name"))
        key = (company.lower(), _date_key(rec))
        groups[key].append(i)

    suffixes = [""] * len(rows)
    for key, indices in groups.items():
        if len(indices) <= 1:
            continue
        # check what differs
        times = [_time_key(rows[i]) for i in indices]
        courses = [_course_key(rows[i]) for i in indices]
        locs = [_loc_key(rows[i]) for i in indices]

        if len(set(courses)) > 1:
            for i in indices:
                suffixes[i] = _course_key(rows[i])
        elif len(set(locs)) > 1:
            for i in indices:
                suffixes[i] = _loc_key(rows[i])
        elif len(set(times)) > 1:
            session = 0
            for i in indices:
                session += 1
                suffixes[i] = f"Session {session}"
        else:
            session = 0
            for i in indices:
                session += 1
                suffixes[i] = f"Session {session}"
    return suffixes


def main():
    ap = argparse.ArgumentParser(description="Build course rosters from an Excel export.")
    ap.add_argument("export", help="Path to the Excel export (.xlsx)")
    ap.add_argument("--templates", default=".", help="Folder containing the blank templates")
    ap.add_argument("--out", default="rosters", help="Output folder")
    ap.add_argument("--keep-order", action="store_true",
                    help="Preserve export order instead of sorting names A-Z")
    ap.add_argument("--include-cancelled", action="store_true",
                    help="Also generate rosters for courses marked CANCELLED")
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)
    template_map = load_template_map(args.templates)
    if not template_map:
        sys.exit(f"No templates found in {args.templates}")
    print("Templates detected:")
    for t, p in template_map.items():
        print(f"  {t:20s} <- {os.path.basename(p)}")

    rows = read_export(args.export)
    print(f"\n{len(rows)} course(s) in export\n")
    suffixes = _assign_suffixes(rows)
    made, skipped = 0, []
    for i, rec in enumerate(rows):
        status, info, ctype, n = build_one(
            rec, template_map, args.out,
            sort_alpha=not args.keep_order,
            include_cancelled=args.include_cancelled,
            fname_suffix=suffixes[i])
        if status == "ok":
            made += 1
            print(f"  [{ctype:20s}] {n:2d} participants -> {os.path.basename(info)}")
        else:
            skipped.append((rec.get("customer"), info))
            print(f"  [SKIPPED            ] {info}")
    print(f"\nDone: {made} roster(s) created"
          + (f", {len(skipped)} skipped" if skipped else ""))


if __name__ == "__main__":
    main()
