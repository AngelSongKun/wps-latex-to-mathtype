#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Shared docx inspection helpers (stdlib only).

- story_parts / strip_fallback / census: formula counting that mirrors what
  the Word object model actually sees. mc:Fallback branches hold dead
  duplicates of mc:Choice content (modern text boxes ship the same shape
  twice: DrawingML in Choice, VML in Fallback); Word never renders or
  converts the Fallback copy, so it must be stripped before counting or
  every census double-counts text-box formulas.
- EMF/WMF '?' glyph scan: MathType draws a literal "?" for any character its
  MathML->MTEF mapping lacks. The OLE preview metafiles in word/media/
  contain ExtTextOut records with the drawn strings, so scanning them finds
  corrupted equations deterministically - no rendering, no OCR, no sampling.
"""
import re
import struct
import zipfile

# Single source of truth for the LaTeX-text detection pattern. Used by
# latex_to_omml.py (conversion) and precheck.py (census) so the two can
# never drift apart. Order matters: $$ display first, then \[ \], then
# $ inline, then \( \).
LATEX_RE = re.compile(
    r"(?P<disp_dd>\$\$(?P<disp_dd_c>[\s\S]+?)\$\$)"
    r"|(?P<disp_br>\\\[+(?P<disp_br_c>[\s\S]+?)\\\]+)"
    r"|(?P<inl_dd>(?<!\\)\$(?P<inl_dd_c>[^$\n]+?)(?<!\\)\$)"
    r"|(?P<inl_paren>\\\((?P<inl_paren_c>[\s\S]+?)\\\))"
)

# chars that MathType historically could not render as raw codepoints.
# Since 2026-11-05 they are re-written to MathML named entities (&vdash;
# &Gg;) in the conversion VBA and render correctly — the census here is a
# regression signal, not a manual-worklist anymore.
KNOWN_MT_GAPS = {
    "\u22a2": "⊢ (\\vdash) — 已修复：命名实体路径，转换后建议目检",
    "\u226b": "≫ (\\gg) — 已修复：命名实体路径，转换后建议目检",
}

# parts that can host equations, matching the Word object model's stories
# (main body, text boxes, footnotes, endnotes, comments, headers, footers)
STORY_PART_RE = re.compile(
    r"word/(document|footnotes|endnotes|comments|header\d*|footer\d*)\.xml$")

FALLBACK_RE = re.compile(r"<mc:Fallback>.*?</mc:Fallback>", re.S)


def story_parts(z):
    return [n for n in z.namelist()
            if n == "word/document.xml" or STORY_PART_RE.match(n)
            and n != "word/document.xml"]


def strip_fallback(xml_text):
    return FALLBACK_RE.sub("", xml_text)


def census(path):
    """(Equation.DSMT4 count, leftover oMath count) across all story parts,
    with mc:Fallback duplicates stripped to match the Word object model."""
    z = zipfile.ZipFile(path)
    ds = left = 0
    for n in z.namelist():
        if n == "word/document.xml" or STORY_PART_RE.match(n):
            d = strip_fallback(z.read(n).decode("utf-8", errors="replace"))
            ds += len(re.findall(r"Equation\.DSMT4", d))
            left += len(re.findall(r"<m:oMath[ >]", d))
    return ds, left


def media_objects(path):
    """Media files referenced by v:imagedata in story parts, document order.

    Returns [(media_name, host_part)] — each w:object (MathType OLE) embeds
    exactly one v:imagedata, so index i corresponds to the i-th equation.
    """
    z = zipfile.ZipFile(path)
    rels = {}
    for n in z.namelist():
        m = re.match(r"word/_rels/(.+)\.rels$", n)
        if m:
            host = "word/" + m.group(1)
            for rid, target in re.findall(
                    r'Id="([^"]+)"[^>]*Target="([^"]+)"',
                    z.read(n).decode("utf-8", errors="replace")):
                rels[(host, rid)] = target
    out = []
    for n in story_parts(z):
        d = strip_fallback(z.read(n).decode("utf-8", errors="replace"))
        for rid in re.findall(r'<v:imagedata[^>]*r:id="([^"]+)"', d):
            target = rels.get((n, rid), "")
            target = target.lstrip("/")
            if not target.startswith("word/"):
                target = "word/" + target
            out.append((target, n))
    return out


def _emf_texts(data):
    """Strings drawn by EMR_EXTTEXTOUTA/W (83/84) and EMR_SMALLTEXTOUT (108)."""
    texts = []
    n = len(data)
    off = 0
    while off + 8 <= n:
        rtype, size = struct.unpack_from("<II", data, off)
        if size < 8 or off + size > n:
            break
        try:
            if rtype in (83, 84):
                # EMRTEXT follows the 36-byte fixed part:
                # ptlReference(8) nChars(4) offString(4) ...
                nchars, offstr = struct.unpack_from("<II", data, off + 36 + 8)
                if 0 < nchars < 100000 and 0 < offstr < size:
                    raw = data[off + offstr: off + offstr + nchars * (2 if rtype == 84 else 1)]
                    if rtype == 84:
                        texts.append(raw.decode("utf-16-le", errors="replace"))
                    else:
                        for cp in ("gbk", "cp1252"):
                            try:
                                texts.append(raw.decode(cp))
                                break
                            except UnicodeDecodeError:
                                continue
            elif rtype == 108:  # EMR_SMALLTEXTOUT, best effort
                cchars, foptions = struct.unpack_from("<II", data, off + 8 + 8)
                if 0 < cchars < 1000:
                    start = off + 36
                    if foptions & 0x8000:  # ETO_UTF? wide chars flag
                        texts.append(data[start:start + cchars * 2]
                                     .decode("utf-16-le", errors="replace"))
                    else:
                        texts.append(data[start:start + cchars]
                                     .decode("cp1252", errors="replace"))
        except struct.error:
            pass
        if rtype == 0:  # EOF
            break
        off += size
    return texts


def _wmf_texts(data):
    """Strings drawn by META_EXTTEXTOUT (0x0A32); 16-bit records in words."""
    texts = []
    if data[:4] == b"\xd7\xcd\xc6\x9a":  # placeable header
        data = data[22:]
    n = len(data)
    off = 18  # skip WMF header
    while off + 6 <= n:
        size_words = struct.unpack_from("<I", data, off)[0]
        func = struct.unpack_from("<H", data, off + 4)[0]
        if size_words < 3 or off + size_words * 2 > n:
            break
        rec = data[off: off + size_words * 2]
        if func == 0x0A32:  # META_EXTTEXTOUT
            try:
                y, x, count, options = struct.unpack_from("<hhHH", rec, 6)
                pos = 14
                if options & 0x0006:  # rect present (ETO_OPAQUE|ETO_CLIPPED)
                    pos += 8
                raw = rec[pos: pos + count]
                for cp in ("gbk", "cp1252"):
                    try:
                        texts.append(raw.decode(cp))
                        break
                    except UnicodeDecodeError:
                        continue
            except struct.error:
                pass
        off += size_words * 2
    return texts


def question_scan(path):
    """Scan every EMF/WMF preview in the docx for literal '?' glyphs.

    Returns {media_name: [text, ...]} — media whose drawn strings contain '?'.
    An empty dict means no equation shows the corruption glyph.
    """
    z = zipfile.ZipFile(path)
    bad = {}
    for n in z.namelist():
        if not re.match(r"word/media/.*\.(emf|wmf)$", n):
            continue
        data = z.read(n)
        try:
            texts = _emf_texts(data) if n.endswith(".emf") else _wmf_texts(data)
        except Exception:
            texts = []
        hits = [t for t in texts if "?" in t]
        if hits:
            bad[n] = hits
    return bad


def equation_previews(path):
    """[(ordinal, media_name, drawn_text)] per MathType w:object, doc order.

    drawn_text is the concatenation of all ExtTextOut strings in the object's
    preview metafile — empty for a blank (failed-set) equation.
    """
    z = zipfile.ZipFile(path)
    out = []
    for i, (mname, _host) in enumerate(media_objects(path)):
        try:
            data = z.read(mname)
            texts = _emf_texts(data) if mname.endswith(".emf") else _wmf_texts(data)
        except Exception:
            texts = []
        out.append((i, mname, "".join(texts)))
    return out


def pdf_artifact_scan(pdf_path):
    """Scan a (Word-exported) PDF for rendering artifacts of MathType objects.

    Detects, per page:
      - question-mark glyphs (U+003F) drawn in LATIN fonts — MathType
        equations must never contain a literal '?'; every one is a font
        substitution artifact (e.g. the NBSP-in-missing-font bug, fixed
        2026-11-05, which drew '?' ON TOP of real glyphs)
      - stacked glyphs: two glyphs sharing the same rounded origin — the
        visual "overlap" signature
    LibreOffice-rendered PDFs cannot be used for this check: its font
    substitution differs from Word's, so it hides/shows different artifacts.
    """
    import collections
    try:
        import pymupdf
    except ImportError:
        return {"error": "pymupdf not installed"}
    doc = pymupdf.open(pdf_path)
    qmarks = []          # (page, font, x, y) — substitution artifacts
    prose_q = 0          # literal '?' in running text (e.g. a citation title)
    stacked = []         # (page, chars, x, y)
    for pno, page in enumerate(doc):
        d = page.get_text("rawdict")
        page_chars = []   # (char, x, y, font, size)
        for b in d["blocks"]:
            if b.get("type") != 0:
                continue
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        x, y = c["origin"]
                        page_chars.append((c["c"], x, y, s["font"], s["size"]))
        for ch, x, y, font, size in page_chars:
            if ch != "?" or "SimSun" in font or "SimHei" in font:
                continue
            # prose '?': the nearest glyph to the LEFT on the same line is a
            # letter ("forecasting?[C]" in references) — legit text
            left = [pc for pc in page_chars
                    if abs(pc[2] - y) < 3 and 0 < x - pc[1] < 10]
            if left and max(left, key=lambda p: p[1])[0].isalpha():
                prose_q += 1
            else:
                qmarks.append((pno + 1, font, round(x), round(y)))
        byorigin = collections.defaultdict(list)
        for ch, x, y, font, size in page_chars:
            byorigin[(round(x), round(y))].append((ch, font, size))
        for (x, y), lst in byorigin.items():
            # real substitution stacking: two DIFFERENT visible glyphs in
            # the SAME font+size at one origin (e.g. '1?'). Narrow glyphs
            # of different fonts (prime ′ + ∈) legitimately share an origin
            # after rounding — skip those.
            distinct = {c for c, _, _ in lst if not c.isspace()}
            same_font = len({(f, round(sz)) for _, f, sz in lst}) == 1
            if len(distinct) >= 2 and same_font:
                stacked.append((pno + 1, "".join(sorted(distinct)), x, y))
    return {"question_marks": qmarks, "prose_question_marks": prose_q,
            "stacked": stacked}
