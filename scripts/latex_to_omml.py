#!/usr/bin/env python3
"""Convert LaTeX formulas (as plain text) in a .docx file to native OMML equations.

Works on documents produced by WPS Office or Microsoft Word (.docx = OOXML).
Scans every paragraph (body, tables, text boxes, footnotes, endnotes,
headers/footers) for LaTeX formulas written as plain text:

    inline:  $...$   or  \\(...\\)
    display: $$...$$  or  \\[...\\]

and replaces them with native Word/WPS equations (OMML), which MathType's
"Convert Equations" batch command can then turn into MathType OLE objects.

Usage:
    python latex_to_omml.py input.docx                    # convert -> input_converted.docx
    python latex_to_omml.py input.docx -o output.docx     # explicit output
    python latex_to_omml.py input.docx --scan             # report only, no changes
    python latex_to_omml.py input.docx --all              # also convert low-confidence matches
    python latex_to_omml.py input.docx --report out.json  # write JSON report

A JSON report is always written next to the output file (or --report path).
"""

import argparse
import copy
import json
import re
import sys
import zipfile
from pathlib import Path

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import os as _os
_sys_dir = _os.path.dirname(_os.path.abspath(__file__))
if _sys_dir not in sys.path:
    sys.path.insert(0, _sys_dir)
from docx_scan import LATEX_RE  # noqa: E402  (single source of truth)

try:
    import latex2mathml.converter
    import mathml2omml
    from docx import Document
    from docx.oxml.ns import qn
    from lxml import etree
except ImportError as e:
    sys.exit(f"Missing dependency: {e}\nInstall with: pip install latex2mathml mathml2omml python-docx lxml")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W_NS, "m": M_NS}


XMLNS_WRAP = f'xmlns:w="{W_NS}" xmlns:m="{M_NS}"'


def confidence(content: str) -> str:
    """Heuristic: is this delimited text really a LaTeX formula?

    Guards against money-style false positives like "价格 $5 和 $10" (CJK
    inside the delimiters without any LaTeX markup). Plain ASCII content
    (``m``, ``N(t)``, ``a+b``) is a typical math variable -> medium.
    """
    if re.search(r"\\[a-zA-Z]+", content):
        return "high"
    if "^" in content or "_" in content:
        return "high"
    if re.search(r"[\u4e00-\u9fff]", content):  # CJK inside $..$ without commands
        return "low"
    return "medium"


# mathml2omml emits a WRONG closing tag for groupChrPr deterministically:
#   <m:groupChr><m:groupChrPr>...</m:groupChr><m:e>...
# (should be </m:groupChrPr>) — one malformed byte stream kills the whole
# document parse, so \bar \vec \widetilde \overbrace \underbrace
# \overrightarrow \check ... all failed. Repair the exact spot: the first
# </m:groupChr> after an unclosed <m:groupChrPr> is that element's close.
_GROUPCHR_BUG = re.compile(
    r"(<m:groupChrPr>(?:(?!</m:groupChrPr>|</m:groupChr>)[\s\S])*?)</m:groupChr>")

# latex2mathml mishandles alignment environments: \begin{aligned} leaves a
# raw '&' (-> invalid XML), \begin{align} silently injects fake equation
# numbers "(1)" into the content. Rewrite them to \begin{gather} (supported,
# multi-line) with alignment marks stripped — layout degrades from aligned
# to centered lines, content survives.
_ALIGN_ENV = re.compile(
    r"\\begin\{(aligned|align\*?|split|flalign\*?|eqnarray\*?)\}"
    r"([\s\S]*?)\\end\{\1\}")


def _preprocess_align_envs(latex: str) -> str:
    def repl(m):
        return "\\begin{gather}" + m.group(2).replace("&", "") + "\\end{gather}"
    return _ALIGN_ENV.sub(repl, latex)


def latex_to_omml_element(latex: str):
    """LaTeX string -> lxml <m:oMath> element. Returns (element, None) or (None, error)."""
    candidates = [latex]
    if _ALIGN_ENV.search(latex):
        candidates.insert(0, _preprocess_align_envs(latex))
    err = None
    for cand in candidates:
        try:
            mathml = latex2mathml.converter.convert(cand.strip())
            omml_str = mathml2omml.convert(mathml)
            omml_str = _GROUPCHR_BUG.sub(r"\1</m:groupChrPr>", omml_str)
            root = etree.fromstring(f"<w:wrap {XMLNS_WRAP}>{omml_str}</w:wrap>")
        except Exception as e:  # bad latex, unsupported macro...
            err = f"{type(e).__name__}: {e}"
            continue
        omath = root[0] if len(root) else None
        if omath is None:
            err = "empty OMML result"
            continue
        return omath, None
    return None, err


def wrap_display_para(omath) -> etree._Element:
    """Wrap an m:oMath in m:oMathPara (standalone centered display equation)."""
    para = etree.fromstring(
        f'<m:oMathPara {XMLNS_WRAP}><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr></m:oMathPara>'
    )
    para.append(omath)
    return para


def make_text_run(text: str, rpr_template) -> etree._Element:
    run = etree.Element(qn("w:r"))
    if rpr_template is not None:
        run.append(copy.deepcopy(rpr_template))
    t = etree.SubElement(run, qn("w:t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return run


def find_template_rpr(p_el):
    for r in p_el.findall(".//w:r", NS):
        rpr = r.find("w:rPr", NS)
        if rpr is not None:
            return rpr
    return None


def in_fallback(p_el) -> bool:
    """True for paragraphs inside mc:Fallback — dead VML duplicates of
    mc:Choice content that Word never renders or exposes as a story;
    converting them would double-convert and pollute the report."""
    return any(etree.QName(a).localname == "Fallback"
               for a in p_el.iterancestors())


def is_container_para(p_el) -> bool:
    """True for anchor paragraphs holding a nested w:p (VML textbox /
    drawing anchor). Rebuilding such a paragraph would destroy the nested
    structure; its inner w:p elements are processed directly instead."""
    return p_el.find(".//w:p", NS) is not None


def process_paragraph(p_el, para_label: str, min_conf: str) -> list:
    """Rewrite LaTeX text in one <w:p> SURGICALLY: only the runs that hold a
    matched span are replaced; every other element (footnote references,
    drawings, fields, bookmarks) stays exactly where it was. The previous
    rebuild-the-whole-paragraph approach destroyed such runs. Returns
    report items."""
    wt_nodes = p_el.findall(".//w:t", NS)
    full = "".join(t.text or "" for t in wt_nodes)
    if "$" not in full and "\\[" not in full and "\\(" not in full:
        return []

    matches = list(LATEX_RE.finditer(full))
    if not matches:
        return []

    preview = re.sub(r"\s+", " ", full).strip()[:40]
    items = []

    # guards: tracked changes / hyperlinks — leave untouched, report
    if p_el.find(".//w:ins", NS) is not None or p_el.find(".//w:hyperlink", NS) is not None:
        for m in matches:
            latex, kind = _match_latex(m)
            items.append(dict(where=para_label, preview=preview, kind=kind,
                              latex=latex.strip(), confidence=confidence(latex),
                              status="skipped",
                              note="paragraph contains hyperlink/tracked changes; convert manually"))
        return items

    order = {"low": 0, "medium": 1, "high": 2}
    only_formula = (len(matches) == 1
                    and not full[:matches[0].start()].strip()
                    and not full[matches[0].end():].strip())

    # cumulative offsets of every w:t inside `full`
    spans = []
    pos = 0
    for t in wt_nodes:
        length = len(t.text or "")
        spans.append([t, pos, pos + length])
        pos += length

    # process matches in REVERSE so earlier offsets stay valid
    for m in reversed(matches):
        latex, kind = _match_latex(m)
        latex = latex.strip()
        conf = confidence(latex)
        if order[conf] < order[min_conf]:
            items.append(dict(where=para_label, preview=preview, kind=kind,
                              latex=latex, confidence=conf, status="skipped",
                              note="low confidence (raise with --all)"))
            continue
        el, err = latex_to_omml_element(latex)
        if el is None:
            items.append(dict(where=para_label, preview=preview, kind=kind,
                              latex=latex, confidence=conf, status="failed",
                              note=err))
            continue

        affected = [sp for sp in spans if sp[2] > m.start() and sp[1] < m.end()]
        first_t, first_s0, _ = affected[0]
        last_t, _, last_s1 = affected[-1]
        before = (first_t.text or "")[: m.start() - first_s0]
        after = (last_t.text or "")[m.end() - last_s1:]

        # affected runs: first (keep `before`), middles (drop), last (keep `after`)
        run_first = first_t.getparent()
        parent = run_first.getparent()
        if parent is None:
            items.append(dict(where=para_label, preview=preview, kind=kind,
                              latex=latex, confidence=conf, status="failed",
                              note="run has no parent"))
            continue
        rpr = run_first.find("w:rPr", NS)
        insert_at = list(parent).index(run_first)

        standalone = only_formula and kind == "display"
        math_el = wrap_display_para(el) if standalone else el

        repl = []
        if before:
            repl.append(make_text_run(before, rpr))
        repl.append(math_el)
        # oMathPara must be a direct child of w:p; if the matched run sits
        # deeper (shouldn't happen — hyperlinks are guarded), fall back inline
        if standalone and parent.tag != etree.QName(p_el).tag:
            repl[-1] = el

        # trim/remove affected text nodes and their runs
        for t, s0, s1 in affected:
            run = t.getparent()
            if run is run_first:
                t.text = before or None
                if not before:
                    run.getparent().remove(run) if run.getparent() is not None else None
            elif run is last_t.getparent():
                t.text = after or None
                if not after and run.getparent() is not None:
                    run.getparent().remove(run)
            else:
                if run.getparent() is not None:
                    run.getparent().remove(run)

        for k, new_el in enumerate(repl):
            parent.insert(insert_at + k, new_el)

        items.append(dict(where=para_label, preview=preview, kind=kind,
                          latex=latex, confidence=conf, status="converted"))
    return items


def _match_latex(m):
    if m.group("disp_dd") is not None:
        return m.group("disp_dd_c"), "display"
    if m.group("disp_br") is not None:
        return m.group("disp_br_c"), "display"
    if m.group("inl_dd") is not None:
        return m.group("inl_dd_c"), "inline"
    return m.group("inl_paren_c"), "inline"


def iter_paragraph_roots(doc):
    """Yield (lxml-root, label) for body, headers and footers."""
    yield doc.element.body, "body"
    for i, sec in enumerate(doc.sections):
        for name in ("header", "footer", "first_page_header", "first_page_footer",
                     "even_page_header", "even_page_footer"):
            part = getattr(sec, name, None)
            if part is None or getattr(part, "is_linked_to_previous", False):
                continue
            yield part._element, f"section{i+1}.{name}"


def iter_footnote_roots(doc):
    """Yield (lxml-root, label, opc-part) for footnotes and endnotes.

    python-docx does not model these parts; reach them through the document
    relationships and patch their blob in place after conversion.
    """
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    for rel in doc.part.rels.values():
        for rel_type, label in ((RT.FOOTNOTES, "footnotes"), (RT.ENDNOTES, "endnotes")):
            if rel.reltype != rel_type or rel.is_external:
                continue
            try:
                root = etree.fromstring(rel.target_part.blob)
            except etree.XMLSyntaxError:
                continue
            yield root, label, rel.target_part


def count_omath(path: Path) -> int:
    from docx_scan import strip_fallback
    with zipfile.ZipFile(path) as z:
        n = 0
        for name in z.namelist():
            if name.startswith("word/") and name.endswith(".xml") and (
                name == "word/document.xml" or "header" in name or "footer" in name
                or "footnotes" in name or "endnotes" in name
                or "comments" in name
            ):
                data = strip_fallback(z.read(name).decode("utf-8", errors="replace"))
                n += data.count("<m:oMath>") + data.count("<m:oMath ")
        return n


def main():
    ap = argparse.ArgumentParser(description="Convert LaTeX text in .docx to OMML equations")
    ap.add_argument("input", help="input .docx file")
    ap.add_argument("-o", "--output", help="output .docx (default: <input>_converted.docx)")
    ap.add_argument("--scan", action="store_true", help="only report candidates, do not modify")
    ap.add_argument("--all", action="store_true", help="convert low-confidence matches too")
    ap.add_argument("--only-high", action="store_true", help="convert only high-confidence matches")
    ap.add_argument("--report", help="path for the JSON report")
    args = ap.parse_args()

    src = Path(args.input)
    if not src.exists():
        sys.exit(f"File not found: {src}")
    min_conf = "low" if args.all else ("high" if args.only_high else "medium")

    doc = Document(str(src))
    items = []
    for root, label in iter_paragraph_roots(doc):
        for i, p_el in enumerate(root.findall(".//w:p", NS), 1):
            if in_fallback(p_el) or is_container_para(p_el):
                continue
            items.extend(process_paragraph(p_el, f"{label}¶{i}", min_conf))

    # footnotes/endnotes: convert on a detached tree, then patch the part
    # blob back so doc.save() serializes it
    for root, label, part in iter_footnote_roots(doc):
        n_before = len(items)
        for i, p_el in enumerate(root.findall(".//w:p", NS), 1):
            if in_fallback(p_el) or is_container_para(p_el):
                continue
            items.extend(process_paragraph(p_el, f"{label}¶{i}", min_conf))
        if any(it["status"] == "converted" for it in items[n_before:]):
            part._blob = etree.tostring(
                root, xml_declaration=True, standalone=True, encoding="UTF-8")

    converted = [it for it in items if it["status"] == "converted"]
    failed = [it for it in items if it["status"] == "failed"]
    skipped = [it for it in items if it["status"] == "skipped"]

    out_path = None
    if not args.scan and converted:
        out_path = Path(args.output) if args.output else src.with_name(src.stem + "_converted.docx")
        doc.save(str(out_path))

    report = {
        "input": str(src),
        "output": str(out_path) if out_path else None,
        "scan_only": args.scan,
        "min_confidence": min_conf,
        "candidates": len(items),
        "converted": len(converted),
        "failed": len(failed),
        "skipped": len(skipped),
        "omath_in_output": count_omath(out_path) if out_path else None,
        "items": items,
    }
    # --scan leaves no files behind unless --report is given explicitly.
    report_path = None
    if args.report or not args.scan:
        report_path = Path(args.report) if args.report else (
            (out_path or src).with_suffix(".report.json"))
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"File: {src.name}")
    print(f"  candidates: {len(items)}  converted: {len(converted)}  "
          f"failed: {len(failed)}  skipped: {len(skipped)}")
    if out_path:
        print(f"  saved to:   {out_path}  ({report['omath_in_output']} native equations)")
    if failed:
        print("\nFailed LaTeX (kept as text):")
        for it in failed:
            print(f"  [{it['where']}] {it['latex'][:60]} -- {it['note']}")
    if skipped:
        print("\nSkipped:")
        for it in skipped:
            print(f"  [{it['where']}] ({it['confidence']}) {it['latex'][:60]} -- {it['note']}")
    if report_path:
        print(f"\nReport: {report_path}")


if __name__ == "__main__":
    main()
