#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""LaTeX symbol <-> MathType coverage audit.

The full conversion chain is:
    LaTeX text --latex2mathml--> MathML --mathml2omml--> OMML
      --OMML2MML.XSL--> MathML --MTSetEqnFromLangStr--> MathType OLE

Failures ("?" glyphs, dropped content, omissions) can happen at any hop.
This tool audits each hop:

  level 1 (default, no Word needed):
      every symbol in MATRIX is run through latex2mathml+mathml2omml.
      Status per symbol: OK / CONVERT_ERROR / EMPTY. Emitted Unicode chars
      are recorded — that is the "LaTeX -> MathML char" half of the mapping.
      Results saved to a JSON file.

  --docx OUT.docx:
      build the audit document: for every level-1-OK symbol one paragraph
      "<latex>  ->  <OMML equation>", grouped by category. Feed this to
      convert_via_word_vba.py, then run --final on the converted file.

  --final CONVERTED.docx --level1 JSON:
      join level-1 results with the converted document: census check +
      per-equation EMF preview '?' scan (MathType draws a literal "?" for
      any character it cannot map). Produces the per-symbol verdict used
      for references/latex-mathtype-map.md.

  --md OUT.md: write the final mapping table as Markdown (with --final).
"""
import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).parent))
from docx_scan import census, equation_previews, question_scan  # noqa: E402

# ---------------------------------------------------------------------------
# Symbol matrix: (category, latex). Everything a thesis realistically uses.
# ---------------------------------------------------------------------------
MATRIX = [
    ("希腊字母(小)", ["\\alpha", "\\beta", "\\gamma", "\\delta", "\\epsilon",
                     "\\varepsilon", "\\zeta", "\\eta", "\\theta", "\\vartheta",
                     "\\iota", "\\kappa", "\\lambda", "\\mu", "\\nu", "\\xi",
                     "\\pi", "\\rho", "\\varrho", "\\sigma", "\\varsigma",
                     "\\tau", "\\upsilon", "\\phi", "\\varphi", "\\chi",
                     "\\psi", "\\omega"]),
    ("希腊字母(大)", ["\\Gamma", "\\Delta", "\\Theta", "\\Lambda", "\\Xi",
                     "\\Pi", "\\Sigma", "\\Upsilon", "\\Phi", "\\Psi", "\\Omega"]),
    ("运算符", ["\\sum", "\\prod", "\\coprod", "\\int", "\\oint", "\\iint",
                "\\iiint", "\\bigcup", "\\bigcap", "\\bigoplus", "\\bigotimes",
                "\\bigvee", "\\bigwedge", "\\lim", "\\limsup", "\\liminf",
                "\\sup", "\\inf", "\\max", "\\min", "\\log", "\\ln", "\\lg",
                "\\exp", "\\sin", "\\cos", "\\tan", "\\cot", "\\sec", "\\csc",
                "\\arcsin", "\\arccos", "\\arctan", "\\sinh", "\\cosh",
                "\\tanh", "\\det", "\\dim", "\\deg", "\\gcd", "\\Pr",
                "\\bmod", "\\pmod{m}", "\\arg", "\\hom", "\\ker"]),
    ("关系符", ["\\leq", "\\geq", "\\neq", "\\leqslant", "\\geqslant", "\\leqq",
                "\\approx", "\\equiv", "\\sim", "\\simeq", "\\cong", "\\propto",
                "\\ll", "\\gg", "\\prec", "\\succ", "\\subseteq", "\\supseteq",
                "\\subset", "\\supset", "\\in", "\\notin", "\\ni", "\\vdash",
                "\\models", "\\perp", "\\parallel", "\\mid", "\\nmid", "\\asymp",
                "\\doteq", "\\colon"]),
    ("箭头", ["\\rightarrow", "\\leftarrow", "\\leftrightarrow", "\\Rightarrow",
              "\\Leftarrow", "\\Leftrightarrow", "\\mapsto", "\\hookrightarrow",
              "\\rightharpoonup", "\\rightleftharpoons", "\\Longrightarrow",
              "\\longrightarrow", "\\uparrow", "\\downarrow", "\\updownarrow",
              "\\nearrow", "\\searrow", "\\leadsto"]),
    ("集合与逻辑", ["\\cup", "\\cap", "\\setminus", "\\emptyset", "\\varnothing",
                    "\\forall", "\\exists", "\\nexists", "\\neg", "\\land",
                    "\\lor", "\\oplus", "\\ominus", "\\otimes", "\\oslash",
                    "\\odot", "\\cdot", "\\circ", "\\bullet", "\\star", "\\ast",
                    "\\dagger", "\\ddagger", "\\pm", "\\mp", "\\times", "\\div",
                    "\\wedge", "\\vee", "\\triangle", "\\diamond", "\\sqcap",
                    "\\sqcup", "\\amalg"]),
    ("重音与装饰", ["\\hat{x}", "\\bar{x}", "\\tilde{x}", "\\dot{x}", "\\ddot{x}",
                    "\\dddot{x}", "\\vec{x}", "\\overline{xy}", "\\underline{x}",
                    "\\widehat{xy}", "\\widetilde{xy}", "\\overbrace{x+y}^{n}",
                    "\\underbrace{x+y}_{n}", "\\overset{a}{b}",
                    "\\underset{a}{b}", "\\stackrel{a}{b}", "\\overleftarrow{AB}",
                    "\\overrightarrow{AB}", "\\breve{a}", "\\check{a}",
                    "\\acute{a}", "\\grave{a}", "\\mathring{a}"]),
    ("定界符", ["\\left(x\\right)", "\\left[x\\right]", "\\left\\{x\\right\\}",
                "\\left|x\\right|", "\\left\\|x\\right\\|", "\\langle x \\rangle",
                "\\lceil x \\rceil", "\\lfloor x \\rfloor", "\\big(x\\big)",
                "\\Big(x\\Big)", "\\bigg(x\\bigg)", "\\Bigg(x\\Bigg)",
                "\\lvert x \\rvert", "\\lVert x \\rVert", "\\vert", "\\Vert"]),
    ("字体与字母", ["\\mathbb{R}", "\\mathbb{N}", "\\mathbb{Z}", "\\mathbb{C}",
                    "\\mathbb{E}", "\\mathbb{Q}", "\\mathbb{P}", "\\mathbb{H}",
                    "\\mathbb{1}", "\\mathcal{F}", "\\mathcal{L}", "\\mathcal{O}",
                    "\\mathfrak{g}", "\\mathfrak{sl}", "\\mathrm{d}x",
                    "\\mathbf{v}", "\\mathit{x}", "\\mathsf{s}", "\\mathtt{t}",
                    "\\hbar", "\\ell", "\\Re", "\\Im", "\\aleph", "\\wp",
                    "\\partial", "\\nabla", "\\imath", "\\jmath", "\\infty",
                    "\\angle", "\\surd", "\\top", "\\bot", "\\forall",
                    "\\flat", "\\natural", "\\sharp"]),
    ("点号", ["\\ldots", "\\cdots", "\\vdots", "\\ddots", "\\dots"]),
    ("结构", ["\\frac{a}{b}", "\\tfrac{a}{b}", "\\dfrac{a}{b}",
              "x^{a}_{b}", "\\binom{n}{k}", "\\sqrt{x}", "\\sqrt[3]{x}",
              "\\sum_{i=1}^{n} x_i", "\\int_{0}^{1} f(t)\\,dt",
              "\\lim_{x \\to 0} f(x)", "\\left. \\frac{dy}{dx} \\right|_{0}^{1}",
              "\\substack{i=1 \\\\ j=2}", "\\overset{\\text{def}}{=}"]),
    ("矩阵与多行", ["\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}",
                    "\\begin{bmatrix} a & b \\\\ c & d \\end{bmatrix}",
                    "\\begin{vmatrix} a & b \\\\ c & d \\end{vmatrix}",
                    "\\begin{matrix} a & b \\\\ c & d \\end{matrix}",
                    "\\begin{cases} x & x>0 \\\\ -x & x\\leq 0 \\end{cases}",
                    "\\begin{array}{cc} a & b \\\\ c & d \\end{array}",
                    "\\begin{aligned} a &= 1 \\\\ b &= 2 \\end{aligned}",
                    "\\begin{align} a &= 1 \\\\ b &= 2 \\end{align}",
                    "\\begin{equation} y = kx + b \\end{equation}",
                    "\\begin{gather} a \\\\ b \\end{gather}",
                    "\\begin{multline} a + b \\\\ + c \\end{multline}",
                    "\\begin{split} a &= 1 \\\\ &= 2 \\end{split}"]),
    ("论文常用组合", ["\\mathbb{E}[X]", "\\hat{Y}_t", "\\bar{R}", "\\|x\\|_2",
                     "\\operatorname{TopK}", "\\operatorname{softmax}",
                     "\\operatorname{ReLU}", "\\mathrm{SmoothL1}",
                     "\\mathcal{L}_{kg}", "\\sigma\\left(\\cdot\\right)",
                     "\\lambda_{max}", "\\bm{v}", "\\bm{\\alpha}",
                     "x^{(l+1)}", "\\mathbf{H}^{(l)}", "\\boldsymbol{\\mu}",
                     "\\text{if } x>0", "\\mbox{OK}", "\\quad", "\\qquad",
                     "\\;x\\;", "\\!x", "\\hspace{2em}x", "\\1", "\\#"]),
]

CHARNAME = {0x2212: "数学负号 MINUS", 0x2032: "prime", 0x2026: "⋯",
            0x22C0: "⋀", 0x22C1: "⋁", 0x22C2: "⋂", 0x22C3: "⋃"}


def omml_chars(omath):
    """Unicode chars inside m:t of an OMML element."""
    import lxml.etree as etree
    M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
    return "".join(t.text or "" for t in omath.iter(f"{M}t"))


def level1():
    try:
        import latex2mathml.converter  # noqa: F401 (dep presence check)
        import mathml2omml  # noqa: F401
    except ImportError as e:
        sys.exit(f"missing dep: {e}")

    from latex_to_omml import latex_to_omml_element  # repaired pipeline
    results = []
    for cat, items in MATRIX:
        for latex in items:
            rec = {"cat": cat, "latex": latex, "status": None, "chars": "",
                   "note": ""}
            el, err = latex_to_omml_element(latex)
            if el is None:
                rec["status"] = "CONVERT_ERROR"
                rec["note"] = str(err)[:100]
            else:
                rec["status"] = "OK"
                rec["chars"] = omml_chars(el)
            results.append(rec)
    return results


def build_docx(results, out_path):
    """Audit document: label text + OMML equation per level-1-OK symbol."""
    from docx import Document
    from docx.oxml.ns import qn
    doc = Document()
    n = 0
    for cat, _ in [(r["cat"], 0) for r in results]:
        pass
    seen_cat = set()
    for rec in results:
        if rec["status"] != "OK":
            continue
        if rec["cat"] not in seen_cat:
            seen_cat.add(rec["cat"])
            h = doc.add_paragraph()
            run = h.add_run(f"### {rec['cat']}")
            run.bold = True
        p = doc.add_paragraph()
        p.add_run(f"[{rec['latex']}] = ")
        from latex_to_omml import latex_to_omml_element
        el, err = latex_to_omml_element(rec["latex"])
        if el is None:
            continue
        p._p.append(el)
        n += 1
    doc.save(out_path)
    return n


def final_report(converted, level1_json):
    l1 = json.loads(Path(level1_json).read_text(encoding="utf-8"))
    ok_syms = [r for r in l1 if r["status"] == "OK"]
    ds, left = census(converted)
    previews = equation_previews(converted)
    bad_media = question_scan(converted)
    for r in ok_syms:
        r["mathtype"] = "ok"
    for i, mname, drawn in previews:
        if i >= len(ok_syms):
            continue
        if mname in bad_media:
            ok_syms[i]["mathtype"] = "MATHTYPE-?(问号)"
        elif not drawn.strip() and ok_syms[i]["chars"].strip():
            # blank preview but the OMML had real content -> lost content
            ok_syms[i]["mathtype"] = "MATHTYPE-FAIL(空白)"
    return {"census": (ds, left), "expected": len(ok_syms),
            "symbols": ok_syms, "bad_media": bad_media}


def write_md(report, out_md):
    lines = ["# LaTeX ↔ MathType 符号对照审计", "",
             f"转换后对象数: {report['census'][0]} / 期望 {report['expected']}"
             f"（残留原生公式 {report['census'][1]}）", "",
             "| LaTeX | 输出字符 | MathType | 备注 |", "|---|---|---|---|"]
    for r in report["symbols"]:
        mark = "✅" if r["mathtype"] == "ok" else "❌ 问号"
        note = r.get("note", "")
        chars = r["chars"]
        if len(chars) > 24:
            chars = chars[:24] + "…"
        lines.append(f"| `{r['latex']}` | `{chars}` | {mark} | {note} |")
    Path(out_md).write_text("\n".join(lines), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description="LaTeX↔MathType 符号审计")
    ap.add_argument("--level1-json", default="symbol_audit_level1.json")
    ap.add_argument("--docx", metavar="OUT.docx", default=None)
    ap.add_argument("--final", metavar="CONVERTED.docx", default=None)
    ap.add_argument("--md", metavar="OUT.md", default=None)
    args = ap.parse_args()

    results = level1()
    Path(args.level1_json).write_text(
        json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8")

    ok = sum(1 for r in results if r["status"] == "OK")
    err = [r for r in results if r["status"] != "OK"]
    print(f"level1: {ok}/{len(results)} OK")
    for r in err:
        print(f"  [{r['status']}] {r['cat']} {r['latex']} -- {r['note']}")

    if args.docx:
        n = build_docx(results, args.docx)
        print(f"audit docx: {n} equations -> {args.docx}")

    if args.final:
        rep = final_report(args.final, args.level1_json)
        print(f"final: converted {rep['census'][0]} expected {rep['expected']} "
              f"left {rep['census'][1]} ?glyph-files {len(rep['bad_media'])}")
        bad = [r for r in rep["symbols"] if r["mathtype"] != "ok"]
        for r in bad:
            print(f"  {r['mathtype']}: {r['latex']} (chars={r['chars']!r})")
        print(f"MATHTYPE-GAP: {len(bad)} 个")
        if args.md:
            write_md(rep, args.md)
            print("mapping table:", args.md)


if __name__ == "__main__":
    main()
