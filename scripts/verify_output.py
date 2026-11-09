#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Post-conversion verification of a MathType-converted .docx.

    python verify_output.py 原始.docx 转换后.docx [--render] [--outdir DIR] [--max-pages 10]

Checks (all stories, not just document.xml):
  1. structural  - every story part of the output has zero leftover OMML;
     total Equation.DSMT4 == word/embeddings/*.bin count == the ORIGINAL's
     oMath census (from precheck.scan) -> nothing dropped, nothing missed.
  2. openable    - python-docx can open the output.
  3. --render    - LibreOffice headless -> PDF, then renders the
     *mandatory-check pages*: pages holding the original's charset-suspect
     formulas (non-cp936 in m:t). Fixed-interval sampling has proven to
     miss these; charset drives the page list instead. PNG filenames are
     unique per run (timestamped) because the Read->CDN upload path serves
     stale cached images when a filename is reused.

Exit code 0 = all checks pass, 1 = failure. Dependencies: pymupdf +
LibreOffice only needed for --render.
"""
import argparse
import os
import re
import subprocess
import sys
import time
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from precheck import scan  # noqa: E402  (stdlib-only)
from docx_scan import census, equation_previews, pdf_artifact_scan, question_scan  # noqa: E402

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]


def find_soffice():
    for p in SOFFICE_CANDIDATES:
        if os.path.isfile(p):
            return p
    return None


WINWORD = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
MT_DIRS = [r"C:\Program Files (x86)\MathType\MathPage\64",
           r"C:\Program Files (x86)\MathType\System\64"]


def export_word_pdf(docx_path, pdf_path):
    """Export via Word itself (the ONLY renderer that reproduces what users
    see/print): LibreOffice substitutes fonts differently and hides artifacts.
    Raises SystemExit-style message on failure."""
    import subprocess
    import time as _t
    import pythoncom
    import win32com.client as wc

    def retry(fn, *a, **k):
        last = None
        for _ in range(60):
            try:
                return fn(*a, **k)
            except pythoncom.com_error as e:
                if e.hresult in (-2147418111, -2147417846):
                    last = e
                    _t.sleep(0.5)
                    continue
                raise
        raise last

    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE"],
                         capture_output=True).stdout
    if b"WINWORD.EXE" in out:
        return ("SKIP: WINWORD.EXE 正在运行（会劫持 COM 附着），"
                "关闭 Word 后重试或手动导出 PDF 后用 --pdf 检查")
    env = dict(os.environ)
    env["PATH"] = os.pathsep.join(MT_DIRS + [env.get("PATH", "")])
    proc = subprocess.Popen([WINWORD, "/embedding"], env=env)
    word = None
    for _ in range(40):
        _t.sleep(1)
        try:
            pythoncom.CoInitialize()
            word = wc.GetActiveObject("Word.Application")
            break
        except pythoncom.com_error:
            pass
    if word is None:
        proc.kill()
        return "SKIP: 无法附着 Word"
    try:
        _t.sleep(3)
        retry(setattr, word, "Visible", True)
        retry(setattr, word, "DisplayAlerts", 0)
        doc = retry(word.Documents.Open, os.path.abspath(docx_path),
                    ReadOnly=True, AddToRecentFiles=False)
        retry(doc.ExportAsFixedFormat, os.path.abspath(pdf_path), 17)
        doc.Close(False)
    finally:
        try:
            retry(word.Quit)
        except Exception:
            pass
        for _ in range(10):
            if proc.poll() is not None:
                break
            _t.sleep(1)
        if proc.poll() is None:
            proc.kill()
    return None


def docx_xml_parts(z):
    return [n for n in z.namelist()
            if n == "word/document.xml"
            or re.match(r"word/(footnotes|endnotes|header\d*|footer\d*)\.xml$", n)]


def main():
    ap = argparse.ArgumentParser(description="MathType 转换结果验证")
    ap.add_argument("input", help="转换前的原始 .docx")
    ap.add_argument("output", help="转换后的 .docx")
    ap.add_argument("--render", action="store_true",
                    help="渲染必检页为 PNG（需 LibreOffice + pymupdf）")
    ap.add_argument("--outdir", default="verify_pages",
                    help="--render 时 PNG/PDF 输出目录（默认 verify_pages/）")
    ap.add_argument("--max-pages", type=int, default=10,
                    help="最多渲染多少个必检页（默认 10）")
    ap.add_argument("--word-pdf", action="store_true",
                    help="用 Word 本体导出 PDF 并扫描渲染工件（'?'字形/同位叠字，"
                         "需本机 Word+pywin32；这是唯一可信的显示层检查）")
    ap.add_argument("--pdf", metavar="EXISTING.pdf", default=None,
                    help="跳过导出，直接扫描一份已有的 Word 导出 PDF")
    args = ap.parse_args()

    problems = []
    src = scan(args.input)

    # 1+2. object census across every story part (mc:Fallback dead duplicates
    # stripped, matching the Word object model)
    out_ds, out_left = census(args.output)
    z = zipfile.ZipFile(args.output)
    emb = len([n for n in z.namelist()
               if n.startswith("word/embeddings/") and n.endswith(".bin")])
    print(f"原始公式: {src['omath_total']}  |  输出 Equation.DSMT4: {out_ds}"
          f"  |  embeddings .bin: {emb}  |  残留 OMML: {out_left}")
    if out_left:
        problems.append(f"残留原生公式 {out_left} 个")
    if out_ds != src["omath_total"]:
        problems.append(f"DSMT4 数 ({out_ds}) != 原始公式数 ({src['omath_total']})")
    if emb != out_ds:
        problems.append(f"embeddings 数 ({emb}) != DSMT4 数 ({out_ds})")

    # 2b. EMF preview '?' scan — MathType draws a literal '?' for any char
    # its mapping lacks; deterministic corruption detector, no sampling
    bad = question_scan(args.output)
    if bad:
        for m, hits in bad.items():
            problems.append(f"EMF 预览含 '?' 字形: {m} -> {hits[:2]}")
    else:
        print("EMF '?' 字形扫描: 0（无 MathType 映射缺口）")

    # 2c. zero-text previews: a blank or bitmap-only preview on a formula
    # that should show content — report for visual follow-up (bitmap-only
    # previews can be legitimate, so warn, don't fail)
    previews = equation_previews(args.output)
    empties = [(i, m) for i, m, t in previews if not t.strip()]
    if empties:
        print(f"零文本预览 {len(empties)} 个（空白或位图绘制，需人工目检）: "              + ", ".join(f"#{i}" for i, _ in empties[:10]))
    else:
        print("零文本预览: 0")

    # 3. python-docx can open it
    try:
        import docx
        docx.Document(args.output)
        print("python-docx 打开: OK")
    except Exception as e:
        problems.append(f"python-docx 打开失败: {e}")

    # 3b. Word-rendered PDF artifact scan ('?' substitution glyphs and
    # same-origin stacked glyphs). This is the ONLY display-layer check that
    # reproduces what users see in Word/print — LibreOffice rendering
    # substitutes fonts differently and hides this class of bug.
    if args.word_pdf or args.pdf:
        if args.pdf:
            wpdf = args.pdf
            skip = None
        else:
            os.makedirs(args.outdir, exist_ok=True)
            wpdf = os.path.join(args.outdir, os.path.basename(
                os.path.splitext(args.output)[0]) + "_word.pdf")
            if os.path.isfile(wpdf):
                os.remove(wpdf)
            skip = export_word_pdf(args.output, wpdf)
        if skip:
            print(skip)
        elif not os.path.isfile(wpdf):
            problems.append("Word PDF 导出失败")
        else:
            art = pdf_artifact_scan(wpdf)
            if "error" in art:
                problems.append(f"PDF 工件扫描失败: {art['error']}")
            else:
                nq, ns = len(art["question_marks"]), len(art["stacked"])
                nprose = art.get("prose_question_marks", 0)
                print(f"Word-PDF 渲染工件: '?'替代字形 {nq} 个, 同位叠字 {ns} 处"
                      f"（应均为 0；另 {nprose} 个正文原生'?'已排除）")
                if nq:
                    pages = sorted({p for p, *_ in art["question_marks"]})
                    problems.append(f"PDF 中 {nq} 个'?'替代字形（页 {pages[:8]}）"
                                    "——公式字体缺字形的替换产物")
                if ns:
                    for p, chars, x, y in art["stacked"][:5]:
                        problems.append(f"第{p}页 同位叠字 {chars!r} @({x},{y})")
                    if ns > 5:
                        problems.append(f"…共 {ns} 处叠字")

    # 4. mandatory-check pages (charset-driven), optional render
    contexts = [f["context"] for f in src["charset"]["formulas"] if f["context"]]
    if args.render:
        soffice = find_soffice()
        if soffice is None:
            problems.append("--render 需要 LibreOffice（未找到 soffice.exe）")
        else:
            os.makedirs(args.outdir, exist_ok=True)
            pdf_path = os.path.join(args.outdir, os.path.basename(
                os.path.splitext(args.output)[0]) + ".pdf")
            # always re-convert: a leftover PDF from an earlier conversion of a
            # same-named file would render stale pages
            if os.path.isfile(pdf_path):
                os.remove(pdf_path)
            subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                            os.path.abspath(args.output),
                            "--outdir", os.path.abspath(args.outdir)],
                           capture_output=True, timeout=300)
            if not os.path.isfile(pdf_path):
                problems.append(f"PDF 转换失败: {pdf_path}")
            else:
                import pymupdf
                doc = pymupdf.open(pdf_path)
                norm = [re.sub(r"\s+", "", p.get_text()) for p in doc]
                pages = []
                for key in contexts:
                    k = re.sub(r"\s+", "", key)
                    hit = next((i for i, t in enumerate(norm) if k in t), None)
                    if hit is not None and hit + 1 not in pages:
                        pages.append(hit + 1)
                pages = sorted(pages)[:args.max_pages]
                stamp = time.strftime("%H%M%S")
                for pg in pages:
                    doc[pg - 1].get_pixmap(dpi=131).save(
                        os.path.join(args.outdir, f"vfy_{stamp}_p{pg:03d}.png"))
                print(f"必检页（字符集驱动, {len(contexts)} 条线索 → "
                      f"{len(pages)} 页）: {pages}")
                print(f"PNG 已输出到 {args.outdir}/（文件名含时间戳，"
                      "避免 CDN 缓存命中旧图）")
    elif contexts:
        print(f"提示: {len(contexts)} 个字符集嫌疑公式未渲染检查（--render 可启用）")

    print("== RESULT ==")
    if problems:
        for p in problems:
            print("FAIL:", p)
        sys.exit(1)
    print("PASS: 结构完整，公式数一致，无残留。请再对必检页做人工/视觉抽查。")
    sys.exit(0)


if __name__ == "__main__":
    main()
