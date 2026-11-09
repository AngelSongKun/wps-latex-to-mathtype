#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Pre-flight scan of a .docx before MathType conversion.

One command answers everything the operator needs before step 2/3:

  1. formula census  - native OMML count per story part (body, text boxes,
     footnotes, endnotes, headers/footers). Text-box/footnote formulas are
     invisible to `doc.OMaths`, which is exactly the "some formulas were
     never recognized" trap; the converter now walks doc.Stories, so this
     census is what the RESULT ok= count should match.
  2. LaTeX residue   - plain-text $...$ / $$...$$ / \\(...\\) / \\[...\\]
     candidates that still need scripts/latex_to_omml.py first.
  3. risky constructs- matrices (m:m), equation arrays (m:eqArr), groupChr,
     borderBox, phantom: OMML2MML.XSL + MathType tend to flatten these;
     they must be rebuilt by hand in MathType after conversion.
  4. charset watch   - non-cp936 characters inside <m:t>. The converter
     entity-encodes them automatically (2026-10-24 fix), so they are no
     longer a corruption risk, but their formulas are the mandatory
     verification pages afterwards - fixed-interval page sampling has
     been proven to miss them.

Usage:
    python precheck.py 输入.docx [--json out.json]

Stdlib only.
"""
import argparse
import json
import re
import sys
import zipfile
from html import unescape

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, str(__import__("os").path.dirname(__import__("os").path.abspath(__file__))))
from docx_scan import KNOWN_MT_GAPS, LATEX_RE, STORY_PART_RE, strip_fallback  # noqa: E402


RISKY = [
    (r"<m:m[ >]", "matrix m:m（矩阵，转换后会被扁平化）"),
    (r"<m:eqArr[ >]", "eq-array m:eqArr（方程组/多行对齐）"),
    (r"<m:groupChr[ >]", "groupChr（组符号，如上/下花括号）"),
    (r"<m:borderBox[ >]", "borderBox（边框盒）"),
    (r"<m:phantom[ >]", "phantom（幻影）"),
]

PART_LABEL = {
    "word/document.xml": "正文+表格",
    "word/footnotes.xml": "脚注",
    "word/endnotes.xml": "尾注",
    "word/comments.xml": "批注",
}


def plain_text(frag):
    return unescape(re.sub(r"<[^>]+>", "", frag))


def cjk_key(doc, pos, length=10):
    """Distinctive CJK context right before pos, for locating the page later."""
    back = plain_text(doc[max(0, pos - 4000):pos])
    cjk = re.findall(r"[\u4e00-\u9fff]{4,}", back)
    return cjk[-1][:length] if cjk else ""


def scan(path):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    report = {"file": path, "omath": {}, "latex_candidates": 0,
              "risky": {}, "charset": {"chars": {}, "formulas": []}}

    total = 0
    latex_all = []
    for n in sorted(names):
        if not STORY_PART_RE.match(n):
            continue
        data = strip_fallback(z.read(n).decode("utf-8", errors="replace"))
        cnt = len(re.findall(r"<m:oMath[ >]", data))
        label = PART_LABEL.get(n) or ("页眉/页脚 " + n)
        if cnt:
            report["omath"][label] = cnt
            total += cnt
        if n == "word/document.xml":
            tb = sum(len(re.findall(r"<m:oMath[ >]", t))
                     for t in re.findall(r"<w:txbxContent>.*?</w:txbxContent>",
                                         data, re.S))
            report["omath_textbox"] = tb
        gap_hits = {}
        for ch, note in KNOWN_MT_GAPS.items():
            k = len(re.findall(re.escape(ch), data))
            if k:
                gap_hits[f"{ch} {note}"] = k
        if gap_hits:
            report["mt_gaps"] = gap_hits
        # LaTeX candidates / risky constructs / charset watch run on EVERY
        # part, matching the all-stories scope of the converter
        latex_all += [m.group(0)[:60] for m in LATEX_RE.finditer(data)]
        for rx, label2 in RISKY:
            c = len(re.findall(rx, data))
            if c:
                report["risky"][label2] = report["risky"].get(label2, 0) + c
        for m in re.finditer(r"<m:oMath[ >].*?</m:oMath>", data, re.S):
            body = m.group(0)
            txt = "".join(plain_text(t) for t in
                          re.findall(r"<m:t(?:\s[^>]*)?>.*?</m:t>", body, re.S))
            bad = {}
            for ch in txt:
                try:
                    ch.encode("cp936")
                except UnicodeEncodeError:
                    bad[ch] = bad.get(ch, 0) + 1
            if bad:
                report["charset"]["formulas"].append({
                    "text": txt[:60],
                    "chars": "".join(bad),
                    "context": cjk_key(data, m.start()),
                    "where": label,
                })
                for ch, c in bad.items():
                    report["charset"]["chars"][ch] = \
                        report["charset"]["chars"].get(ch, 0) + c
    report["latex_candidates"] = len(latex_all)
    report["latex_samples"] = latex_all[:5]
    report["omath_total"] = total
    return report


def main():
    ap = argparse.ArgumentParser(description="docx 公式转换预检")
    ap.add_argument("input")
    ap.add_argument("--json", metavar="PATH", default=None)
    args = ap.parse_args()

    r = scan(args.input)
    print(f"File: {args.input}")
    print("== 公式分布（OMML） ==")
    for k, v in r["omath"].items():
        print(f"  {k}: {v}")
    if r.get("omath_textbox"):
        print(f"  （其中文本框内: {r['omath_textbox']} —— 旧版只转正文会漏掉）")
    print(f"  合计: {r['omath_total']}")

    print("== LaTeX 文本候选 ==")
    print(f"  candidates: {r['latex_candidates']}")
    for s in r.get("latex_samples", []):
        print(f"    e.g. {s}")

    if r.get("mt_gaps"):
        print("== 已知 MathType 缺口字符（数据层无法转换，转换后需手工重建） ==")
        for k, v in r["mt_gaps"].items():
            print(f"  {k}  ×{v}")

    print("== 风险结构（转换后可能扁平化，需手工重建） ==")
    print("  " + ("；".join(f"{k}: {v}" for k, v in r["risky"].items())
                  if r["risky"] else "无"))

    cs = r["charset"]
    print(f"== 字符集（{len(cs['formulas'])} 个公式含 GBK 外字符；"
          f"EntityEncode 已自动修复，转换后这些公式所在页为必检页） ==")
    if cs["chars"]:
        print("  " + "  ".join(f"{ch!r}×{c}" for ch, c in
                               sorted(cs["chars"].items(), key=lambda x: -x[1])))
        for f in cs["formulas"][:8]:
            print(f"    [{f['context']}] {f['text']}  ({f['chars']})")
        if len(cs["formulas"]) > 8:
            print(f"    …共 {len(cs['formulas'])} 个")

    print("== 结论 ==")
    if r["latex_candidates"] and not r["omath_total"]:
        print("  纯 LaTeX 文本 → 先跑第 2 步 latex_to_omml.py，再转换")
    elif r["latex_candidates"]:
        print("  OMML + LaTeX 残留混合 → 先跑第 2 步转 LaTeX，再转换")
    elif r["omath_total"]:
        print("  全部为原生公式 → 直接第 3 步 convert_via_word_vba.py"
              f"（预期 ok={r['omath_total']}）")
    else:
        print("  未发现任何公式（无 OMML、无 LaTeX 文本）→ 无需转换")
    if r["risky"]:
        print("  注意：存在风险结构，转换后须按清单逐个手工核对")
    if args.json:
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(r, f, ensure_ascii=False, indent=2)
        print(f"Report: {args.json}")


if __name__ == "__main__":
    main()
