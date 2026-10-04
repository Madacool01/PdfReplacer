"""
Dad's PDF Keyword Replacer - compact classic tkinter UI.
- Select PDFs or folder, find/replace one keyword pair
- Keep original font/size or set custom, plus bold/italic/underline/highlight
- Saves to new folder, originals untouched
- Run: python pdf_replacer.py  |  Build EXE: pyinstaller --onefile --windowed --name DadPdfReplacer pdf_replacer.py
Requires: pymupdf
"""
import os
import csv
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser, ttk
from pathlib import Path

try:
    import pymupdf
except ImportError:
    raise SystemExit("Missing dependency: pip install pymupdf")

# ---------- font helpers ----------

BASE_FONTS = {
    ("helv", False, False): "helv",
    ("helv", True, False): "hebo",
    ("helv", False, True): "heit",
    ("helv", True, True): "hebi",
    ("cour", False, False): "cour",
    ("cour", True, False): "cobo",
    ("cour", False, True): "coit",
    ("cour", True, True): "cobi",
    ("tiro", False, False): "tiro",
    ("tiro", True, False): "tibo",
    ("tiro", False, True): "tiit",
    ("tiro", True, True): "tibi",
}
FAMILY_LABELS = {"helv": "Helvetica", "cour": "Courier", "tiro": "Times"}


def detect_family_and_flags(span_font: str):
    f = (span_font or "").lower()
    if "cour" in f or "courier" in f or "mono" in f:
        fam = "cour"
    elif "time" in f or "tiro" in f or "serif" in f or "georgia" in f:
        fam = "tiro"
    else:
        fam = "helv"
    bold = any(k in f for k in ("bold", "black", "heavy", "demi", "boto", "hebo", "cobo", "tibo"))
    italic = any(k in f for k in ("italic", "oblique", "it ", "heit", "coit", "tiit"))
    return fam, bold, italic


def resolve_fontfile(span_font: str, bold: bool, italic: bool):
    """Best-effort match to a real .ttf so inserted text looks like the original.

    Returns a fontfile path or None (caller falls back to Base14).
    """
    import os as _os
    f = (span_font or "").lower()
    windir = _os.environ.get("WINDIR", r"C:\Windows")
    fd = _os.path.join(windir, "Fonts")

    def pick(*names):
        for n in names:
            p = _os.path.join(fd, n)
            if _os.path.exists(p):
                return p
        return None

    if "tahoma" in f:
        if bold and italic:
            return pick("tahomabd.ttf", "arialbi.ttf")
        if bold:
            return pick("tahomabd.ttf", "arialbd.ttf")
        if italic:
            return pick("tahomait.ttf", "ariali.ttf") if _os.path.exists(_os.path.join(fd, "tahomait.ttf")) else pick("ariali.ttf")
        return pick("tahoma.ttf", "arial.ttf")
    if "univers" in f and "cond" in f:
        # Closest stock Windows match for Univers Condensed is Arial Narrow
        if bold and italic:
            return pick("ARIALNBI.TTF", "arialbi.ttf")
        if bold:
            return pick("ARIALNB.TTF", "arialbd.ttf")
        if italic:
            return pick("ARIALNI.TTF", "ariali.ttf")
        return pick("ARIALN.TTF", "arial.ttf")
    if "arial" in f and "narrow" in f:
        if bold and italic:
            return pick("ARIALNBI.TTF")
        if bold:
            return pick("ARIALNB.TTF")
        if italic:
            return pick("ARIALNI.TTF")
        return pick("ARIALN.TTF")
    if "arial" in f or "helv" in f or "univers" in f or "sans" in f:
        if bold and italic:
            return pick("arialbi.ttf")
        if bold:
            return pick("arialbd.ttf")
        if italic:
            return pick("ariali.ttf")
        return pick("arial.ttf")
    if "time" in f or "tiro" in f or "serif" in f or "georgia" in f:
        if bold and italic:
            return pick("timesbi.ttf")
        if bold:
            return pick("timesbd.ttf")
        if italic:
            return pick("timesi.ttf")
        return pick("times.ttf")
    if "cour" in f or "mono" in f:
        if bold and italic:
            return pick("courbi.ttf")
        if bold:
            return pick("courbd.ttf")
        if italic:
            return pick("couri.ttf")
        return pick("cour.ttf")
    # generic sans fallback
    if bold:
        return pick("arialbd.ttf")
    return pick("arial.ttf")


_sysfont_registry_cache = None
_sysfont_lookup_cache = {}
_pdffont_extract_cache = {}


def _font_family_key(name):
    import re as _re
    return _re.sub(r"[^a-z0-9]+", "", (name or "").lower())


def _read_windows_font_registry():
    """Read HKLM font display-name -> file mapping once (fast, accurate)."""
    global _sysfont_registry_cache
    if _sysfont_registry_cache is not None:
        return _sysfont_registry_cache
    _sysfont_registry_cache = []
    try:
        import winreg as _wr
        _k = _wr.OpenKey(_wr.HKEY_LOCAL_MACHINE,
                         r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
        try:
            _n = _wr.QueryInfoKey(_k)[1]
            for _i in range(_n):
                try:
                    _nm, _vl, _tp = _wr.EnumValue(_k, _i)
                    _sysfont_registry_cache.append((_nm, _vl))
                except OSError:
                    break
        finally:
            _wr.CloseKey(_k)
    except Exception:
        pass
    return _sysfont_registry_cache


def find_system_font(span_font, bold, italic):
    """Full system font (with complete charset) matching the span's family.

    Uses the Windows font registry, so ANY family (Castellar, Tahoma,
    Univers...) resolves -- not just the few hardcoded ones. Returns a
    full path or None.
    """
    import os as _os
    fam = _font_family_key((_os.path.basename(span_font or "").split("+")[-1]).split("-")[0])
    if not fam:
        return None
    key = (fam, bool(bold), bool(italic))
    if key in _sysfont_lookup_cache:
        return _sysfont_lookup_cache[key]
    found = None
    try:
        entries = _read_windows_font_registry()
        cands = []
        for disp, fn in entries:
            dl = disp.lower()
            if fam not in _font_family_key(disp):
                continue
            is_b = "bold" in dl or "demi" in dl or "black" in dl or "heavy" in dl
            is_i = "italic" in dl or "oblique" in dl
            # score: style must match; prefer exact family word + TrueType
            score = 0
            score += 0 if is_b == bool(bold) else 5
            score += 0 if is_i == bool(italic) else 5
            if "(truetype)" in dl:
                score -= 2
            cands.append((score, disp, fn))
        cands.sort(key=lambda t: (t[0], t[1]))
        fd = _os.path.join(_os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
        for _, _, fn in cands:
            p = fn if _os.path.isabs(fn) else _os.path.join(fd, fn)
            if _os.path.exists(p):
                found = p
                break
    except Exception:
        found = None
    _sysfont_lookup_cache[key] = found
    return found


def extract_pdf_font(doc, span_font, need_text=""):
    """Extract the embedded font bytes from the PDF to a temp .ttf.

    Last-resort path when no full system font exists: guarantees the
    exact original outlines. Subsets may miss glyphs for brand-new
    characters, so a glyph-coverage check against need_text decides.
    Returns (path_or_None, covers_all_bool).
    """
    import os as _os
    import tempfile as _tf
    fam = (span_font or "").split("+")[-1]
    ckey = (_font_family_key(fam), _font_family_key(need_text[:64]))
    if ckey in _pdffont_extract_cache:
        return _pdffont_extract_cache[ckey]
    result = (None, False)
    try:
        # locate xref of the font on any page of the doc
        target = None
        for _pg in doc:
            try:
                for _f in _pg.get_fonts(full=True):
                    _xref, _ext = _f[0], _f[1]
                    _base = _f[3] or ""
                    if _font_family_key(_base.split("+")[-1].split("-")[0]) == _font_family_key(fam.split("-")[0]) \
                            or _font_family_key(_base) == _font_family_key(fam):
                        target = _xref
                        break
            except Exception:
                continue
            if target:
                break
        if target:
            info = doc.extract_font(target)
            content = info.get("content") if isinstance(info, dict) else None
            ext = (info.get("ext") if isinstance(info, dict) else "") or "ttf"
            if content:
                dd = _os.path.join(_tf.gettempdir(), "dadpdf_fonts")
                _os.makedirs(dd, exist_ok=True)
                fp = _os.path.join(dd, "pdfembed-%d.%s" % (target, ext))
                if not _os.path.exists(fp):
                    with open(fp, "wb") as _fh:
                        _fh.write(content)
                covers = True
                if need_text:
                    try:
                        _fo = pymupdf.Font(fontfile=fp)
                        for _ch in set(need_text):
                            if _ch in (" ", "\t", "\n"):
                                continue
                            try:
                                if not _fo.has_glyph(ord(_ch)):
                                    covers = False
                                    break
                            except Exception:
                                continue
                    except Exception:
                        covers = False
                result = (fp, covers)
    except Exception:
        result = (None, False)
    _pdffont_extract_cache[ckey] = result
    return result


def resolve_font(doc, span_font, bold, italic, need_text=""):
    """Best fontfile for inserted text: system full font first (complete
    charset, exact family), then embedded PDF font, then legacy mapping."""
    p = find_system_font(span_font, bold, italic)
    if p:
        return p
    try:
        fp, covers = extract_pdf_font(doc, span_font, need_text)
        if fp and covers:
            return fp
        if fp:
            return fp  # partial subset still better than a wrong family
    except Exception:
        pass
    try:
        return resolve_fontfile(span_font, bold, italic)
    except Exception:
        return None


def measure_width(text, fontname, fontsize, fontfile=None):
    if not text:
        return 0.0
    try:
        if fontfile:
            fo = pymupdf.Font(fontfile=fontfile)
            return fo.text_length(text, fontsize=fontsize)
        return pymupdf.get_text_length(text, fontname=fontname, fontsize=fontsize)
    except Exception:
        try:
            return pymupdf.get_text_length(text, fontname=fontname, fontsize=fontsize)
        except Exception:
            return len(text) * fontsize * 0.5


def int_to_rgb(color_int):
    if color_int is None:
        return (0, 0, 0)
    try:
        r = ((color_int >> 16) & 255) / 255
        g = ((color_int >> 8) & 255) / 255
        b = (color_int & 255) / 255
        return (r, g, b)
    except Exception:
        return (0, 0, 0)


def hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def rgb_to_hex(rgb):
    return "#%02x%02x%02x" % tuple(int(max(0, min(1, c)) * 255) for c in rgb)


# ---------- core replace logic ----------
# Line-based replacement: redact the whole line that contains the match and
# rewrite the full line. This fixes the classic "DIA inside DIAPHRAGM" bug
# where only the 3-letter sub-rect was cleared, leaving "PHRAGM" visible
# underneath the new text.

def _find_occurrences(haystack, needle, match_case):
    if not needle:
        return []
    h = haystack if match_case else haystack.lower()
    n = needle if match_case else needle.lower()
    out, start = [], 0
    while True:
        i = h.find(n, start)
        if i < 0:
            return out
        out.append(i)
        start = i + max(1, len(n))


def _occurrence_is_whole_word(haystack, idx, needle_len):
    before = haystack[idx - 1] if idx > 0 else " "
    after = haystack[idx + needle_len] if idx + needle_len < len(haystack) else " "
    return not (before.isalnum() or before == "_" or after.isalnum() or after == "_")


def _sub_line_text(line_text, find_str, replace_str, match_case, whole_word):
    """Return (new_text, n_replacements) for one dict line."""
    new_text, n, _ = _sub_line_text_ex(line_text, find_str, replace_str, match_case, whole_word)
    return new_text, n


def _sub_line_text_ex(line_text, find_str, replace_str, match_case, whole_word):
    """Return (new_text, n_replacements, spans) for one dict line.

    spans = list of (start, end) offsets of each inserted replacement
    inside new_text, used to scope highlight/underline to just the
    new text instead of the whole rewritten line.
    """
    import re as _re
    if not find_str:
        return line_text, 0, []
    if whole_word:
        flags = 0 if match_case else _re.IGNORECASE
        pat = _re.compile(r"(?<![A-Za-z0-9_])" + _re.escape(find_str) + r"(?![A-Za-z0-9_])", flags)
    elif match_case:
        pat = _re.compile(_re.escape(find_str))
    else:
        pat = _re.compile(_re.escape(find_str), _re.IGNORECASE)
    out = []
    spans = []
    last = 0
    n = 0
    for m in pat.finditer(line_text):
        s, e = m.span()
        out.append(line_text[last:s])
        ns = sum(len(p) for p in out)
        out.append(replace_str)
        spans.append((ns, ns + len(replace_str)))
        last = e
        n += 1
    out.append(line_text[last:])
    return "".join(out), n, spans


def _line_direction(line_dir):
    """Normalize a dict line['dir'] vector; fall back to horizontal."""
    try:
        dx, dy = float(line_dir[0]), float(line_dir[1])
    except Exception:
        return (1.0, 0.0)
    import math as _m
    n = _m.hypot(dx, dy)
    if n < 1e-9:
        return (1.0, 0.0)
    return (dx / n, dy / n)


def _rotation_for_direction(dx, dy):
    """Insertion transform following direction (dx, dy).

    Returns (rotate_int_or_None, morph_matrix_or_None). Axis-aligned
    text uses insert_text(rotate=...) (exact round-trip; rotate only
    accepts 0/90/180/270); any other angle uses a morph rotation
    matrix. Dict coords are y-down, so R = atan2(-dy, dx).
    """
    import math as _m
    ang = _m.degrees(_m.atan2(-dy, dx)) % 360.0
    for cand in (0, 90, 180, 270):
        if abs((ang - cand + 180.0) % 360.0 - 180.0) < 2.0:
            return cand, None
    r = _m.radians(ang)
    return None, (round(_m.cos(r), 6), round(_m.sin(r), 6),
                  round(-_m.sin(r), 6), round(_m.cos(r), 6), 0.0, 0.0)


def _embed_fontname(fontfile):
    """Unique reference name for an embedded TTF.

    IMPORTANT: must NOT be a Base14 name (helv/hebo/tiro/...). Tests show
    PyMuPDF silently ignores `fontfile` when `fontname` is a Base14 name
    and embeds Helvetica instead -- the exact font-mismatch bug reported.
    """
    import os as _os
    import re as _re
    stem = _os.path.splitext(_os.path.basename(fontfile or ""))[0]
    stem = _re.sub(r"[^A-Za-z0-9]+", "-", stem).strip("-").lower() or "emb"
    return stem[:32] + "-emb"


def get_span_style(page, rect):
    """Kept for compatibility: find span with max overlap (used as fallback)."""
    try:
        d = page.get_text("dict")
    except Exception:
        return None
    best = None
    best_area = 0
    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                sb = pymupdf.Rect(span["bbox"])
                inter = sb & rect
                if inter.is_empty:
                    continue
                area = abs(inter.get_area())
                if area > best_area:
                    best_area = area
                    best = span
    return best


def matches_filter(clip_text, find_str, match_case, whole_word):
    if not match_case:
        ct, fs = clip_text.lower(), find_str.lower()
    else:
        ct, fs = clip_text, find_str
    idx = ct.find(fs)
    if idx < 0:
        return False
    if whole_word:
        before = ct[idx - 1] if idx > 0 else " "
        after = ct[idx + len(fs)] if idx + len(fs) < len(ct) else " "
        if before.isalnum() or before == "_" or after.isalnum() or after == "_":
            return False
    return True


def find_match_rects(page, find_str, match_case, whole_word):
    """Count-oriented helper: one rect per occurrence (used by Preview).

    Uses line-text scanning so a sub-word hit like DIA in DIAPHRAGM counts.
    """
    try:
        d = page.get_text("dict")
    except Exception:
        try:
            return [pymupdf.Rect(r) for r in page.search_for(find_str)]
        except Exception:
            return []
    out = []
    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            line_text = "".join(s.get("text", "") for s in spans)
            _, n = _sub_line_text(line_text, find_str, "", match_case, whole_word)
            if n <= 0:
                # fall back to search_for for ligature/encoding edge cases
                try:
                    for r in page.search_for(find_str):
                        sb = pymupdf.Rect(line.get("bbox", spans[0]["bbox"]))
                        if sb.intersects(r):
                            out.append(pymupdf.Rect(r))
                except Exception:
                    pass
                continue
            # one rect per occurrence, approximated from the line bbox
            try:
                lb = pymupdf.Rect(line["bbox"])
                w = max(lb.width, 1)
                for k in range(n):
                    x0 = lb.x0 + w * k / max(n, 1)
                    out.append(pymupdf.Rect(x0, lb.y0, x0 + 2, lb.y1))
            except Exception:
                pass
    return out


def collect_line_jobs(page, find_str, replace_str, match_case, whole_word):
    """Scan dict lines; return jobs covering the FULL line bbox.

    Each job: {rect, new_text, count, span(style source), line_bbox}
    """
    try:
        d = page.get_text("dict")
    except Exception:
        return []
    jobs = []
    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            spans = [s for s in line.get("spans", []) if s.get("text")]
            if not spans:
                continue
            line_text = "".join(s.get("text", "") for s in spans)
            new_text, n, rep_spans = _sub_line_text_ex(line_text, find_str, replace_str, match_case, whole_word)
            if n <= 0:
                continue
            # dominant span = longest span containing the match, else longest span
            dom = max(spans, key=lambda s: len(s.get("text", "")))
            occ = _find_occurrences(line_text, find_str, match_case)
            if occ and whole_word:
                occ = [i for i in occ if _occurrence_is_whole_word(line_text, i, len(find_str))]
            if occ:
                # pick span covering first occurrence offset
                pos = occ[0]
                acc = 0
                for s in spans:
                    t = s.get("text", "")
                    if acc <= pos < acc + len(t):
                        dom = s
                        break
                    acc += len(t)
            try:
                x0 = min(s["bbox"][0] for s in spans)
                y0 = min(s["bbox"][1] for s in spans)
                x1 = max(s["bbox"][2] for s in spans)
                y1 = max(s["bbox"][3] for s in spans)
                lb = pymupdf.Rect(x0, y0, x1, y1)
            except Exception:
                continue
            jobs.append({"rect": lb, "new_text": new_text, "count": n, "span": dom,
                         "old_text": line_text, "rep_spans": rep_spans,
                         "dir": _line_direction(line.get("dir", (1, 0)))})
    return jobs


def count_matches(pdf_path, find_str, match_case, whole_word):
    total = 0
    try:
        doc = pymupdf.open(pdf_path)
    except Exception:
        return 0
    with doc:
        for page in doc:
            for j in collect_line_jobs(page, find_str, "", match_case, whole_word):
                total += j["count"]
    return total


def count_bookmark_matches(pdf_path, find_str, match_case, whole_word):
    """Count find_str occurrences inside PDF bookmark (outline) titles."""
    total = 0
    try:
        doc = pymupdf.open(pdf_path)
    except Exception:
        return 0
    with doc:
        try:
            toc = doc.get_toc()
        except Exception:
            return 0
        for entry in toc:
            try:
                _, n, _ = _sub_line_text_ex(entry[1], find_str, "", match_case, whole_word)
            except Exception:
                continue
            total += n
    return total


def count_matches_all(pdf_path, find_str, match_case, whole_word):
    """Return (text_matches, bookmark_matches) for Preview breakdown."""
    return (count_matches(pdf_path, find_str, match_case, whole_word),
            count_bookmark_matches(pdf_path, find_str, match_case, whole_word))


def replace_bookmarks_in_doc(doc, find_str, replace_str, match_case, whole_word):
    """Replace find_str in every bookmark title of an open document.

    Returns the number of replacements. Only touches the outline when
    at least one title matches, so bookmark-free PDFs are unaffected.
    """
    try:
        toc = doc.get_toc()
    except Exception:
        return 0
    if not toc or not find_str:
        return 0
    total = 0
    changed = False
    new_toc = []
    for entry in toc:
        try:
            lvl, title = entry[0], entry[1]
            rest = entry[2:]
        except Exception:
            new_toc.append(entry)
            continue
        try:
            new_title, n, _ = _sub_line_text_ex(title, find_str, replace_str, match_case, whole_word)
        except Exception:
            new_toc.append(entry)
            continue
        if n > 0:
            changed = True
            total += n
            new_toc.append([lvl, new_title] + list(rest))
        else:
            new_toc.append(entry)
    if changed:
        doc.set_toc(new_toc)
    return total


def replace_in_pdf(src, dst, find_str, replace_str, opts):
    """
    opts: dict with match_case, whole_word,
          keep_font / keep_size / keep_color (each bool; fall back to
          legacy keep_original for all three),
          bold, italic, underline, highlight, highlight_hex,
          custom_family(helv/cour/tiro), custom_size, custom_hex,
          replace_text (bool), replace_bookmarks (bool)
    Returns (total_count, warning_or_None, text_count, bookmark_count)
    """
    match_case = opts["match_case"]
    whole_word = opts["whole_word"]
    _legacy_keep = opts.get("keep_original", True)
    keep_font = opts.get("keep_font", _legacy_keep)
    keep_size = opts.get("keep_size", _legacy_keep)
    keep_color = opts.get("keep_color", _legacy_keep)
    do_text = opts.get("replace_text", True)
    do_bookmarks = opts.get("replace_bookmarks", False)

    try:
        doc = pymupdf.open(src)
    except Exception as e:
        return 0, f"Cannot open: {e}", 0, 0
    if doc.is_encrypted:
        try:
            doc.authenticate("")
        except Exception:
            pass
        if doc.is_encrypted:
            doc.close()
            return 0, "Skipped (encrypted/password protected)", 0, 0

    total = 0
    text_total = 0
    bm_total = 0
    if do_bookmarks:
        try:
            bm_total = replace_bookmarks_in_doc(doc, find_str, replace_str, match_case, whole_word)
            total += bm_total
        except Exception:
            pass
    try:
        for page in doc:
            if not do_text:
                break
            raw_jobs = collect_line_jobs(page, find_str, replace_str, match_case, whole_word)
            if not raw_jobs:
                continue
            # gather styles BEFORE redacting (whole-line redact => no leftovers)
            jobs = []
            for rj in raw_jobs:
                rect = rj["rect"]
                span = rj["span"]
                new_text = rj["new_text"]
                rep_spans = sorted(rj.get("rep_spans", []))
                # Original style: untouched parts of the line ALWAYS keep
                # this, so only the replacement visibly changes.
                if span:
                    det_fam, det_b, det_i = detect_family_and_flags(span.get("font", ""))
                    orig_size = float(span.get("size", 11))
                    orig_rgb = int_to_rgb(span.get("color", 0))
                    try:
                        _org = span.get("origin", [rect.x0, rect.y1 - 2])
                        ox, oy = float(_org[0]), float(_org[1])
                        base_y = oy
                    except Exception:
                        ox, oy = rect.x0, rect.y1 - 2
                        base_y = oy
                    orig_font = span.get("font", "")
                else:
                    det_fam, det_b, det_i = opts["custom_family"], opts["bold"], opts["italic"]
                    orig_size, orig_rgb = float(opts["custom_size"]), hex_to_rgb(opts["custom_hex"])
                    ox, oy = rect.x0, rect.y1 - 2
                    base_y = oy
                    orig_font = ""
                # writing direction of the original line: replacements are
                # inserted along it so vertical text stays vertical, etc.
                ddx, ddy = rj.get("dir", (1.0, 0.0)) if span else (1.0, 0.0)
                rot, morphmat = _rotation_for_direction(ddx, ddy)
                # Replacement style: per-attribute original/custom mix.
                # Bold/Italic ticks apply to the replacement only.
                if keep_font and span:
                    rep_fam, rep_b, rep_i = det_fam, det_b, det_i
                else:
                    rep_fam, rep_b, rep_i = opts["custom_family"], opts["bold"], opts["italic"]
                if opts["bold"]:
                    rep_b = True
                if opts["italic"]:
                    rep_i = True
                rep_fontname = BASE_FONTS.get((rep_fam, rep_b, rep_i), "helv")
                if keep_font and span:
                    rep_fontfile = resolve_font(doc, orig_font, rep_b, rep_i, new_text)
                else:
                    rep_fontfile = None
                    # custom font: map chosen family to a real ttf when possible
                    fam_to_probe = {"helv": "arial.ttf", "tiro": "times.ttf", "cour": "cour.ttf"}.get(rep_fam, "")
                    if fam_to_probe:
                        import os as _os2
                        _p = _os2.path.join(_os2.environ.get("WINDIR", r"C:\Windows"), "Fonts", fam_to_probe)
                        if _os2.path.exists(_p):
                            rep_fontfile = _p
                rep_size = orig_size if (keep_size and span) else float(opts["custom_size"])
                rep_rgb = orig_rgb if (keep_color and span) else hex_to_rgb(opts["custom_hex"])
                # Original-segment style (remainder of the line)
                orig_fontname = BASE_FONTS.get((det_fam, det_b, det_i), "helv")
                if span:
                    orig_fontfile = resolve_font(doc, orig_font, det_b, det_i, new_text)
                else:
                    orig_fontfile = rep_fontfile
                # Split the rewritten line: replacement parts get the new
                # style, untouched parts keep the original style.
                segs = []  # (text, is_replacement)
                _pos = 0
                for (_rs, _re) in rep_spans:
                    if _rs > _pos:
                        segs.append((new_text[_pos:_rs], False))
                    if _re > _rs:
                        segs.append((new_text[_rs:_re], True))
                    _pos = max(_pos, _re)
                if _pos < len(new_text):
                    segs.append((new_text[_pos:], False))
                if not segs:
                    segs = [(new_text, True)]
                # shrink-to-fit against the ORIGINAL line LENGTH (width for
                # horizontal, height for vertical) so longer text never
                # spills over neighbouring symbols/words; scale both parts
                # uniformly to keep the line harmonious
                if new_text:
                    try:
                        _wrep = sum(measure_width(t, rep_fontname, rep_size, rep_fontfile) for t, r in segs if r)
                        _worg = sum(measure_width(t, orig_fontname, orig_size, orig_fontfile) for t, r in segs if not r)
                        _linelen = rect.height if abs(ddy) > abs(ddx) else rect.width
                        avail = max(_linelen, 5)
                        if _wrep + _worg > avail + 0.5:
                            _f = avail / max(_wrep + _worg, 0.01)
                            rep_size = max(4.0, rep_size * _f)
                            orig_size = max(4.0, orig_size * _f)
                    except Exception:
                        pass
                eff_rep = _embed_fontname(rep_fontfile) if rep_fontfile else rep_fontname
                eff_orig = _embed_fontname(orig_fontfile) if orig_fontfile else orig_fontname
                seg_objs = []
                for (_t, _isrep) in segs:
                    if _t == "":
                        continue
                    if _isrep:
                        seg_objs.append({"text": _t, "fontname": eff_rep, "fontfile": rep_fontfile,
                                         "size": rep_size, "rgb": rep_rgb, "is_rep": True})
                    else:
                        seg_objs.append({"text": _t, "fontname": eff_orig, "fontfile": orig_fontfile,
                                         "size": orig_size, "rgb": orig_rgb, "is_rep": False})
                jobs.append({"rect": rect, "base_y": base_y, "ox": ox, "oy": oy,
                             "dx": ddx, "dy": ddy, "rotate": rot, "morphmat": morphmat,
                             "new_text": new_text, "count": rj["count"],
                             "rep_spans": rep_spans, "segs": seg_objs,
                             "rep_rgb": rep_rgb,
                             # legacy single-style keys (first segment style)
                             "fontname": eff_rep, "fontfile": rep_fontfile,
                             "size": rep_size, "rgb": rep_rgb})

            # redact whole lines but CLAMPED so a tight label stack
            # (e.g. DIAPHR AGM above PUMP, 0.38pt apart) never deletes the
            # neighbour line: apply_redactions removes any glyph touching
            # the redact rect, so we must not touch the neighbour's bbox.
            try:
                _d = page.get_text("dict")
                _other = []
                for _b in _d.get("blocks", []):
                    for _l in _b.get("lines", []):
                        try:
                            _other.append(pymupdf.Rect(_l["bbox"]))
                        except Exception:
                            pass
            except Exception:
                _other = []
            for j in jobs:
                try:
                    r = pymupdf.Rect(j["rect"])
                    for ob in _other:
                        # skip self / near-identical box
                        if abs(ob.x0 - r.x0) < 0.5 and abs(ob.y0 - r.y0) < 0.5 \
                                and abs(ob.x1 - r.x1) < 0.5 and abs(ob.y1 - r.y1) < 0.5:
                            continue
                        # neighbour below overlaps our bottom edge
                        if ob.y0 > r.y0 and ob.y0 < r.y1:
                            # only if x-ranges overlap (same label stack)
                            if not (ob.x1 < r.x0 or ob.x0 > r.x1):
                                r.y1 = min(r.y1, ob.y0 - 0.35)
                        # neighbour above overlaps our top edge
                        elif ob.y1 < r.y1 and ob.y1 > r.y0:
                            if not (ob.x1 < r.x0 or ob.x0 > r.x1):
                                r.y0 = max(r.y0, ob.y1 + 0.35)
                    if r.y1 - r.y0 < 1:
                        r.y1 = r.y0 + 1
                    j["safe_rect"] = pymupdf.Rect(r)
                    page.add_redact_annot(pymupdf.Rect(r), fill=(1, 1, 1))
                except Exception:
                    try:
                        j["safe_rect"] = pymupdf.Rect(j["rect"])
                        page.add_redact_annot(pymupdf.Rect(j["rect"]), fill=(1, 1, 1))
                    except Exception:
                        pass
            try:
                page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE)
            except TypeError:
                page.apply_redactions()
            # paint white over the full original line box to hide any sliver
            # left by clamping (pixel paint only, does not delete neighbours)
            for j in jobs:
                try:
                    page.draw_rect(pymupdf.Rect(j["rect"]), color=None, fill=(1, 1, 1),
                                   overlay=True, width=0)
                except Exception:
                    pass

            # insert new line segment by segment from the original origin,
            # advancing along the original writing direction so vertical
            # text stays vertical. Replacement parts use the new style,
            # untouched parts keep the original style (e.g. only PIA red).
            for j in jobs:
                rect = j["rect"]
                ox = min(max(j["ox"], 0), page.rect.x1 - 1)
                oy = min(max(j["oy"], 0), page.rect.y1 - 1)
                dx, dy = j.get("dx", 1.0), j.get("dy", 0.0)
                rot, mmat = j.get("rotate", 0), j.get("morphmat")

                def _scope_rect(off0, seg_w):
                    ax0, ay0 = ox + dx * off0, oy + dy * off0
                    ax1, ay1 = ox + dx * (off0 + seg_w), oy + dy * (off0 + seg_w)
                    rx0, rx1 = min(ax0, ax1) - 1, max(ax0, ax1) + 1
                    ry0, ry1 = min(ay0, ay1) - 1, max(ay0, ay1) + 1
                    if abs(dx) >= abs(dy):
                        ry0, ry1 = rect.y0, rect.y1
                    else:
                        rx0, rx1 = rect.x0, rect.x1
                    return pymupdf.Rect(max(rx0, 0), max(ry0, 0),
                                        min(rx1, page.rect.x1 - 1), min(ry1, page.rect.y1 - 1))

                scope_rects = []
                _off = 0.0
                for _sg in j.get("segs", []):
                    _t = _sg["text"]
                    if _t == "":
                        continue
                    _pt = pymupdf.Point(ox + dx * _off, oy + dy * _off)
                    kwargs = dict(fontname=_sg["fontname"], fontsize=_sg["size"], color=_sg["rgb"])
                    if _sg["fontfile"]:
                        kwargs["fontfile"] = _sg["fontfile"]
                    if mmat is not None:
                        kwargs["morph"] = (_pt, pymupdf.Matrix(*mmat))
                    else:
                        kwargs["rotate"] = rot if rot is not None else 0
                    try:
                        page.insert_text(_pt, _t, **kwargs)
                    except Exception:
                        try:
                            _fb = dict(kwargs)
                            _fb.pop("morph", None)
                            _fb.pop("fontfile", None)
                            page.insert_text(_pt, _t, **_fb)
                        except Exception:
                            page.insert_text(_pt, _t,
                                             fontsize=_sg["size"], color=_sg["rgb"])
                    try:
                        _w = measure_width(_t, _sg["fontname"], _sg["size"], _sg["fontfile"])
                    except Exception:
                        _w = 0
                    if _sg.get("is_rep"):
                        scope_rects.append(_scope_rect(_off, _w))
                    _off += _w
                total += j["count"]
                text_total += j["count"]
                if not scope_rects:
                    # fallback (e.g. deletion with highlight on): whole line
                    try:
                        w = measure_width(j["new_text"], j["fontname"], j["size"], j["fontfile"])
                    except Exception:
                        _linelen = rect.height if abs(dy) > abs(dx) else rect.width
                        w = _linelen
                    scope_rects = [_scope_rect(0.0, w)]
                if opts["highlight"]:
                    hc = hex_to_rgb(opts["highlight_hex"])
                    for new_rect in scope_rects:
                        try:
                            a = page.add_highlight_annot(new_rect)
                            a.set_colors(stroke=hc)
                            a.set_opacity(0.45)
                            a.update()
                        except Exception:
                            pass
                if opts["underline"]:
                    for new_rect in scope_rects:
                        try:
                            a = page.add_underline_annot(new_rect)
                            a.set_colors(stroke=j["rgb"])
                            a.update()
                        except Exception:
                            # fallback: draw a line along the writing direction
                            try:
                                if abs(dy) > abs(dx):
                                    _mx = (new_rect.x0 + new_rect.x1) / 2 + 1
                                    p1 = pymupdf.Point(_mx, new_rect.y0)
                                    p2 = pymupdf.Point(_mx, new_rect.y1)
                                else:
                                    p1 = pymupdf.Point(new_rect.x0, new_rect.y1 - 1)
                                    p2 = pymupdf.Point(new_rect.x1, new_rect.y1 - 1)
                                page.draw_line(p1, p2, color=j["rgb"], width=0.8)
                            except Exception:
                                pass
        # metadata
        try:
            doc.set_metadata({**doc.metadata, "producer": "Dad PDF Replacer"})
        except Exception:
            pass
        Path(dst).parent.mkdir(parents=True, exist_ok=True)
        doc.save(dst, garbage=4, deflate=True)
        doc.close()
        return total, None, text_total, bm_total
    except Exception as e:
        try:
            doc.close()
        except Exception:
            pass
        return total, f"Error: {e}", text_total, bm_total


# ---------- UI ----------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Dad's PDF Keyword Replacer")
        self.geometry("680x780")
        self.files = []

        self.var_find = tk.StringVar()
        self.var_replace = tk.StringVar()
        self.var_case = tk.BooleanVar(value=False)
        self.var_whole = tk.BooleanVar(value=True)
        self.var_scope = tk.StringVar(value="text")
        self.var_keep_font = tk.BooleanVar(value=True)
        self.var_keep_size = tk.BooleanVar(value=True)
        self.var_keep_color = tk.BooleanVar(value=True)
        self.var_bold = tk.BooleanVar(value=False)
        self.var_italic = tk.BooleanVar(value=False)
        self.var_underline = tk.BooleanVar(value=False)
        self.var_highlight = tk.BooleanVar(value=False)
        self.var_family = tk.StringVar(value="Helvetica")
        self.var_size = tk.IntVar(value=11)
        self.var_textcolor = tk.StringVar(value="#000000")
        self.var_hlcolor = tk.StringVar(value="#FFFF00")
        self.var_out = tk.StringVar(value=str(Path.cwd() / "fixed"))

        self.build()

    def build(self):
        pad = {"padx": 10, "pady": 4}
        tk.Label(self, text="1. Choose PDF files", font=("Arial", 12, "bold")).pack(anchor="w", **pad)
        fr = tk.Frame(self)
        fr.pack(fill="x", **pad)
        tk.Button(fr, text="Add PDFs", font=("Arial", 11), command=self.add_files, width=12).pack(side="left", padx=4)
        tk.Button(fr, text="Add Folder", font=("Arial", 11), command=self.add_folder, width=12).pack(side="left", padx=4)
        tk.Button(fr, text="Remove", font=("Arial", 11), command=self.remove_sel).pack(side="left", padx=4)
        tk.Button(fr, text="Clear", font=("Arial", 11), command=self.clear).pack(side="left", padx=4)

        self.listbox = tk.Listbox(self, height=7, font=("Arial", 10), selectmode="extended")
        self.listbox.pack(fill="x", **pad)
        self.lbl_count = tk.Label(self, text="0 files selected", font=("Arial", 10))
        self.lbl_count.pack(anchor="w", padx=12)

        tk.Label(self, text="2. What to find and replace", font=("Arial", 12, "bold")).pack(anchor="w", **pad)
        gf = tk.Frame(self)
        gf.pack(fill="x", **pad)
        tk.Label(gf, text="Find:", font=("Arial", 11), width=12, anchor="e").grid(row=0, column=0)
        tk.Entry(gf, textvariable=self.var_find, font=("Arial", 11), width=40).grid(row=0, column=1, padx=4)
        tk.Label(gf, text="Replace with:", font=("Arial", 11), width=12, anchor="e").grid(row=1, column=0)
        tk.Entry(gf, textvariable=self.var_replace, font=("Arial", 11), width=40).grid(row=1, column=1, padx=4)
        of = tk.Frame(self)
        of.pack(fill="x", padx=10)
        tk.Checkbutton(of, text="Match case", variable=self.var_case, font=("Arial", 10)).pack(side="left")
        tk.Checkbutton(of, text="Whole word only", variable=self.var_whole, font=("Arial", 10)).pack(side="left", padx=12)
        wf = tk.Frame(self)
        wf.pack(fill="x", padx=10, pady=4)
        tk.Label(wf, text="Where:", font=("Arial", 10)).pack(side="left")
        seg = tk.Frame(wf, relief="sunken", bd=1)
        seg.pack(side="left", padx=6)
        for _val, _txt in (("text", "PDF text"), ("bookmarks", "Bookmarks"),
                           ("both", "Text + Bookmarks")):
            tk.Radiobutton(seg, text=_txt, value=_val, variable=self.var_scope,
                           indicatoron=0, width=16, font=("Arial", 10),
                           selectcolor="#c8e6c9").pack(side="left", padx=1, pady=1)

        tk.Label(self, text="3. How should the NEW word look?", font=("Arial", 12, "bold")).pack(anchor="w", **pad)
        sf = tk.Frame(self)
        sf.pack(fill="x", padx=10)
        tk.Checkbutton(sf, text="Original font", variable=self.var_keep_font,
                       font=("Arial", 10), command=self.toggle_custom).pack(side="left")
        tk.Checkbutton(sf, text="Original size", variable=self.var_keep_size,
                       font=("Arial", 10), command=self.toggle_custom).pack(side="left", padx=12)
        tk.Checkbutton(sf, text="Original color", variable=self.var_keep_color,
                       font=("Arial", 10), command=self.toggle_custom).pack(side="left", padx=12)
        self.custom_frame = tk.Frame(self)
        self.custom_frame.pack(fill="x", padx=10)
        tk.Label(self.custom_frame, text="Font:", font=("Arial", 10)).grid(row=0, column=0, sticky="e")
        self.cbo_family = ttk.Combobox(self.custom_frame, textvariable=self.var_family,
                                       values=["Helvetica", "Times", "Courier"], width=12, state="readonly")
        self.cbo_family.grid(row=0, column=1, padx=4)
        tk.Label(self.custom_frame, text="Size:", font=("Arial", 10)).grid(row=0, column=2, sticky="e")
        self.spin_size = tk.Spinbox(self.custom_frame, from_=6, to=72, textvariable=self.var_size, width=5, font=("Arial", 10))
        self.spin_size.grid(row=0, column=3, padx=4)
        self.btn_textcolor = tk.Button(self.custom_frame, text="Text color", command=self.pick_text)
        self.btn_textcolor.grid(row=0, column=4, padx=4)
        self.btn_text_prev = tk.Label(self.custom_frame, text="   ", bg="#000000", width=4)
        self.btn_text_prev.grid(row=0, column=5)

        bf = tk.Frame(self)
        bf.pack(fill="x", padx=10, pady=4)
        tk.Checkbutton(bf, text="Bold", variable=self.var_bold, font=("Arial", 11)).pack(side="left")
        tk.Checkbutton(bf, text="Italic", variable=self.var_italic, font=("Arial", 11)).pack(side="left", padx=8)
        tk.Checkbutton(bf, text="Underline", variable=self.var_underline, font=("Arial", 11)).pack(side="left", padx=8)
        tk.Checkbutton(bf, text="Highlight", variable=self.var_highlight, font=("Arial", 11)).pack(side="left", padx=8)
        tk.Button(bf, text="Highlight color", command=self.pick_hl).pack(side="left", padx=4)
        self.btn_hl_prev = tk.Label(bf, text="   ", bg="#FFFF00", width=4)
        self.btn_hl_prev.pack(side="left")

        tk.Label(self, text="4. Replace", font=("Arial", 12, "bold")).pack(anchor="w", **pad)
        outf = tk.Frame(self)
        outf.pack(fill="x", padx=10)
        tk.Label(outf, text="Save to:", font=("Arial", 10)).pack(side="left")
        tk.Entry(outf, textvariable=self.var_out, font=("Arial", 10), width=42).pack(side="left", padx=4)
        tk.Button(outf, text="Browse", command=self.browse_out).pack(side="left")

        btnf = tk.Frame(self)
        btnf.pack(fill="x", padx=10, pady=8)
        tk.Button(btnf, text="Preview (count only)", font=("Arial", 11), command=self.preview).pack(side="left", padx=4)
        self.btn_go = tk.Button(btnf, text="REPLACE IN ALL FILES", font=("Arial", 12, "bold"),
                                bg="#2e7d32", fg="white", command=self.run)
        self.btn_go.pack(side="left", padx=10)

        self.prog = ttk.Progressbar(self, length=600, mode="determinate")
        self.prog.pack(padx=10, pady=4)
        self.status = tk.Label(self, text="Ready. Pick files to start.", font=("Arial", 10), fg="#333")
        self.status.pack(anchor="w", padx=12)
        self.toggle_custom()

    def toggle_custom(self):
        # each custom control is only editable when its "Original ..." tick is OFF
        try:
            self.cbo_family.configure(state="readonly" if not self.var_keep_font.get() else "disabled")
        except Exception:
            pass
        try:
            self.spin_size.configure(state="normal" if not self.var_keep_size.get() else "disabled")
        except Exception:
            pass
        try:
            self.btn_textcolor.configure(state="normal" if not self.var_keep_color.get() else "disabled")
        except Exception:
            pass

    def pick_text(self):
        c = colorchooser.askcolor(self.var_textcolor.get())[1]
        if c:
            self.var_textcolor.set(c)
            self.btn_text_prev.configure(bg=c)

    def pick_hl(self):
        c = colorchooser.askcolor(self.var_hlcolor.get())[1]
        if c:
            self.var_hlcolor.set(c)
            self.btn_hl_prev.configure(bg=c)
            if not self.var_highlight.get():
                self.var_highlight.set(True)

    def browse_out(self):
        d = filedialog.askdirectory()
        if d:
            self.var_out.set(d)

    def refresh_count(self):
        self.lbl_count.config(text=f"{len(self.files)} files selected")

    def add_files(self):
        ps = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        for p in ps:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert("end", p)
        self.refresh_count()

    def add_folder(self):
        d = filedialog.askdirectory()
        if not d:
            return
        for p in sorted(Path(d).glob("*.pdf")):
            s = str(p)
            if s not in self.files:
                self.files.append(s)
                self.listbox.insert("end", s)
        self.refresh_count()

    def remove_sel(self):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)
            del self.files[i]
        self.refresh_count()

    def clear(self):
        self.files.clear()
        self.listbox.delete(0, "end")
        self.refresh_count()

    def get_opts(self):
        fam_map = {"Helvetica": "helv", "Times": "tiro", "Courier": "cour"}
        kf, ks, kc = self.var_keep_font.get(), self.var_keep_size.get(), self.var_keep_color.get()
        scope = self.var_scope.get()
        return {
            "match_case": self.var_case.get(),
            "whole_word": self.var_whole.get(),
            "replace_text": scope in ("text", "both"),
            "replace_bookmarks": scope in ("bookmarks", "both"),
            "keep_original": bool(kf and ks and kc),  # legacy compat
            "keep_font": kf,
            "keep_size": ks,
            "keep_color": kc,
            "bold": self.var_bold.get(),
            "italic": self.var_italic.get(),
            "underline": self.var_underline.get(),
            "highlight": self.var_highlight.get(),
            "highlight_hex": self.var_hlcolor.get(),
            "custom_family": fam_map.get(self.var_family.get(), "helv"),
            "custom_size": int(self.var_size.get()),
            "custom_hex": self.var_textcolor.get(),
        }

    def validate(self):
        if not self.files:
            messagebox.showwarning("No files", "Please add PDF files first.")
            return False
        if not self.var_find.get():
            messagebox.showwarning("Missing keyword", "Please type the keyword to FIND.")
            return False
        if self.var_replace.get() == "":
            if not messagebox.askyesno("Empty replacement", "Replace-with is empty. This will DELETE the keyword. Continue?"):
                return False
        return True

    def preview(self):
        if not self.validate():
            return
        find = self.var_find.get()
        opts = self.get_opts()
        t_total, b_total = 0, 0
        per_file = []
        self.status.config(text="Counting...")
        self.update()
        for f in self.files:
            t, b = count_matches_all(f, find, opts["match_case"], opts["whole_word"])
            if not opts["replace_text"]:
                t = 0
            if not opts["replace_bookmarks"]:
                b = 0
            per_file.append((os.path.basename(f), t, b))
            t_total += t
            b_total += b
        total = t_total + b_total
        msg = f"Found {total} matches in {len(self.files)} files ({t_total} in text, {b_total} in bookmarks).\n\n"
        msg += "\n".join(f"{n}: {t} text, {b} bookmarks" for n, t, b in per_file[:20])
        if len(per_file) > 20:
            msg += f"\n...and {len(per_file) - 20} more files"
        messagebox.showinfo("Preview", msg)
        self.status.config(text=f"Preview: {total} matches found ({t_total} text, {b_total} bookmarks).")

    def run(self):
        if not self.validate():
            return
        find = self.var_find.get()
        repl = self.var_replace.get()
        opts = self.get_opts()
        outdir = Path(self.var_out.get())
        outdir.mkdir(parents=True, exist_ok=True)

        self.prog["maximum"] = len(self.files)
        self.prog["value"] = 0
        self.btn_go.config(state="disabled")
        grand = 0
        grand_text = 0
        grand_bm = 0
        rows = []
        for i, f in enumerate(self.files):
            self.status.config(text=f"Working {i + 1}/{len(self.files)}: {os.path.basename(f)}")
            self.update()
            dst = outdir / Path(f).name
            # avoid overwriting source if same folder: add _fixed suffix
            try:
                if Path(f).resolve() == dst.resolve():
                    dst = outdir / (Path(f).stem + "_fixed.pdf")
            except Exception:
                pass
            n, warn, tn, bn = replace_in_pdf(f, str(dst), find, repl, opts)
            grand += n
            grand_text += tn
            grand_bm += bn
            rows.append([f, str(dst), tn, bn, n, warn or "OK"])
            self.prog["value"] = i + 1
            self.update()

        log = outdir / "replace_log.csv"
        with open(log, "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["source", "output", "text_replacements", "bookmark_replacements", "total", "status"])
            w.writerows(rows)
        self.btn_go.config(state="normal")
        self.status.config(text=f"Done! {grand} replacements ({grand_text} text, {grand_bm} bookmarks) in {len(self.files)} files. Log: {log}")
        messagebox.showinfo("Done", f"Done!\n{grand} replacements ({grand_text} text, {grand_bm} bookmarks) in {len(self.files)} files.\nSaved to: {outdir}\nLog: replace_log.csv")


if __name__ == "__main__":
    App().mainloop()
