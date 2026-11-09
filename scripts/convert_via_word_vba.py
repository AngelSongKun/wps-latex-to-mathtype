# -*- coding: utf-8 -*-
"""OMML -> MathType OLE converter: one-time macro add-in, zero-touch runs.

Usage:
    python convert_via_word_vba.py input.docx [more.docx ...] [-o out.docx]
        [--kill-word] [--limit N] [--rebuild-addin]

- converts native equations (OMML) in EVERY story — body, tables, text boxes,
  footnotes, endnotes, headers/footers — to MathType Equation.DSMT4 objects
- the input file itself is never written to; the working document lives at the
  output path and is checkpoint-saved every SAVE_EVERY formulas
- RESUME: a partial output (converted DSMT4 + leftover oMath == input count)
  continues from its checkpoint automatically — including after a Word CRASH:
  the script relaunches Word and resumes without manual intervention
- --limit N converts only the first N formulas (smoke testing)
- BATCH: several input files are converted in sequence, each in its own
  Word session but sharing one process launch per file
- MACRO ADD-IN (2026-11-05): the conversion macro lives in a versioned
  PyMathTypeConv_vN.dotm installed ONCE into the Word Startup folder
  (the same mechanism MathType itself uses). Normal runs need NO VBE access
  at all — VBE automation was the flakiest part of the old pipeline
  (E_FAIL races costing minutes of restart cycles). Bump ADDIN_VER whenever
  VBA_CODE changes; the stale add-in is rebuilt automatically. Delete the
  .dotm from the Startup folder to uninstall.
- engine: MTInitAPI / MTXFormReset / MTSetEqnFromLangStr / MTCloseOleObject
  from MathPage.WLL (late-bound declares, correct export names)
- OMML->MathML: late-bound MSXML2.DOMDocument30 + OMML2MML.XSL (ProgID
  registration per fix_environment.py)
- MathML cleanups before the engine (see troubleshooting.md):
    NBSP -> plain space (preview font lacks F0A0 -> '?' stacked on glyphs)
    U+00B7 middle dot -> U+22C5 dot operator (lone 0xB7 = GBK lead byte)
    non-ASCII -> numeric entities (&#8722;) — survive every parser
- accents: precomposed/combining accent chars are rewritten to m:acc in the
  docx before Word opens it (MathType cannot draw them; m:acc renders
  perfectly everywhere) — decompose_accents()
- insertion replicates MathType's InsertMTEqnFromMML: InsertFile(BlankEqn.doc,
  "MTBlankEqn") -> DoVerb(2) -> MTSetEqnFromLangStr -> MTCloseOleObject
  (must use rng.InsertFile + rng.Expand wdWord + rng.InlineShapes(1);
  Selection.InlineShapes does NOT see the freshly inserted object)

A WINWORD.EXE left running from a crashed run (or the user's own session)
breaks the GetActiveObject attach step, so the script refuses to start while
any Word process is alive unless --kill-word is given. Stale ~$ owner-lock
files next to the input are removed automatically.
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import unicodedata
import zipfile

import pythoncom
import win32com.client as wc
from lxml import etree

sys_dir = os.path.dirname(os.path.abspath(__file__))
if sys_dir not in sys.path:
    sys.path.insert(0, sys_dir)
from docx_scan import census  # noqa: E402

WINWORD = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
MT = r"C:\Program Files (x86)\MathType"

ADDIN_VER = 3  # bump whenever VBA_CODE changes -> stale add-in rebuilt
ADDIN_NAME = "PyMathTypeConv_v%d.dotm" % ADDIN_VER
STARTUP_DIR = os.path.join(os.environ.get("APPDATA", ""),
                           "Microsoft", "Word", "STARTUP")
ADDIN_PATH = os.path.join(STARTUP_DIR, ADDIN_NAME)

VBA_CODE = r'''
Option Explicit

Private Declare PtrSafe Function MTInitAPI Lib "MathPage.WLL" (ByVal options As Integer, ByVal timeout As Integer) As Long
Private Declare PtrSafe Function MTXFormReset Lib "MathPage.WLL" () As Long
Private Declare PtrSafe Function MTSetEqnFromLangStr Lib "MathPage.WLL" (ByVal myObj As Object, ByVal langType As Integer, langStr As Any, ByVal strLen As Long) As Long
Private Declare PtrSafe Function MTCloseOleObject Lib "MathPage.WLL" (ByVal dwSaveOpts As Long, ByVal varOleObject As Object) As Long

Private Const XSL_PATH As String = "C:\Program Files\Microsoft Office\root\Office16\OMML2MML.XSL"
Private Const BLANK_EQN As String = "C:\Program Files (x86)\MathType\Office Support\BlankEqn.doc"
Private Const SAVE_EVERY As Long = 20

' Iterate doc.Stories, not doc.OMaths: the latter only covers the main text
' story and silently skips text boxes, footnotes, endnotes, headers/footers.
' testLimit (passed by the Python driver as Application.Run argument) is the
' --limit smoke-test cap; 0 means unlimited.
Public Function PyConvertAll(Optional testLimit As Long = 0) As String
    On Error GoTo eh
    Dim r As Long
    r = MTInitAPI(0, 30)
    If r <= 0 Then PyConvertAll = "FAIL init=" & r: Exit Function
    r = MTXFormReset()
    If r <> 0 Then PyConvertAll = "FAIL xformreset=" & r: Exit Function

    Dim doc As Document
    Set doc = ActiveDocument

    Dim style As Object, src As Object
    Set style = CreateObject("MSXML2.DOMDocument30")
    If Not style.Load(XSL_PATH) Then
        PyConvertAll = "FAIL xsl-load: " & style.parseError.reason
        Exit Function
    End If
    Set src = CreateObject("MSXML2.DOMDocument30")

    Dim ok As Long, fail As Long, n0 As Long, i As Long
    Dim statHist As String
    Dim leftAll As Long

    ' snapshot every story range first (Document.OMaths only covers the main
    ' text story; text boxes/footnotes/endnotes/headers need StoryRanges,
    ' chained via NextStoryRange across sections). Modifying content while
    ' enumerating the live collection is unsafe, hence the snapshot.
    Dim storyList As New Collection
    Dim sr As Range, cur As Range
    For Each sr In doc.StoryRanges
        Set cur = sr
        Do While Not cur Is Nothing
            storyList.Add cur
            Set cur = cur.NextStoryRange
        Loop
    Next sr

    Dim k As Long
    For k = 1 To storyList.Count
        Dim story As Range
        Set story = storyList(k)
        n0 = story.OMaths.Count
        If n0 > 0 Then
        i = n0

        Do While i >= 1
            Dim om As OMath, rng As Range, wo As String, mml As String
            Set om = story.OMaths(i)
            Set rng = om.Range
            On Error Resume Next
            Err.Clear

        wo = rng.WordOpenXML
        src.LoadXML wo
        mml = src.transformNode(style.documentElement)
        ' OMML2MML.XSL turns every plain space in m:t into an <mi> holding
        ' U+00A0 (NBSP). MathType stores it as a character and the preview
        ' draws byte 0xA0 — which several preview fonts lack (modern
        ' symbol.ttf has no F0A0), so Word's renderer substitutes a literal
        ' '?' that OVERLAPS the adjacent glyph in the exported PDF
        ' (2026-11-05: 47 stray '?', 7 stacked, across the thesis). A plain
        ' space exists in every font and is dropped or drawn invisibly.
        mml = Replace(mml, ChrW(160), ChrW(32))
        ' U+00B7 middle dot: MathType groups it into CJK-styled (SimSun/GBK)
        ' text runs when the formula also holds Chinese words, and a lone
        ' 0xB7 byte is a GBK lead byte -> Word renders '?' (thesis 2026-11-05,
        ' M_loc mask formula). U+22C5 dot operator renders via Symbol font
        ' in every context (verified by the 301-symbol audit) and is the
        ' proper math multiplication dot anyway.
        mml = Replace(mml, ChrW(183), ChrW(8901))
        ' VARIANT-TEST v2: named MathML entities for chars MathType
        ' refuses as raw codepoints
        mml = Replace(mml, ChrW(8866), "&vdash;")
        mml = Replace(mml, ChrW(8811), "&Gg;")
        ' V3: <mi> semantics so lone-operator formulas are not dropped
        mml = Replace(mml, "<mo>&vdash;</mo>", "<mi>&vdash;</mi>")
        mml = Replace(mml, "<mo>&Gg;</mo>", "<mi>&Gg;</mi>")
        mml = EntityEncode(mml)

            Dim res As Boolean, stat As Long
            res = False
            stat = -12345
            If Err.Number = 0 And Len(mml) > 0 Then
                res = InsertMT(rng, mml, stat)
            End If

            If Err.Number <> 0 Then
                fail = fail + 1
                doc.UndoClear
            ElseIf res Then
                ok = ok + 1
            Else
                fail = fail + 1
                Dim tag As String
                ' ordinal lets the caller map a failed set back to the
                ' equation: processing runs LAST->FIRST, so ordinal k is
                ' document position (total - k + 1)
                tag = "|" & stat & "@" & (ok + fail)
                If InStr(statHist, Left$(tag, InStr(tag & "@", "@"))) = 0 Then statHist = statHist & tag
            End If
            On Error GoTo eh
            i = i - 1
            If testLimit > 0 And (ok + fail) >= testLimit Then Exit Do
            If (ok + fail) Mod 10 = 0 Then
                Application.StatusBar = "PyConv " & (ok + fail) & " ok=" & ok & " fail=" & fail
                DoEvents
            End If
            ' periodic in-place save = crash checkpoint (doc already lives at
            ' the output path after the immediate SaveAs2 from Python)
            If SAVE_EVERY > 0 And (ok + fail) Mod SAVE_EVERY = 0 Then doc.Save
        Loop
        End If
    Next k

    For Each sr In doc.StoryRanges
        leftAll = leftAll + sr.OMaths.Count
        Dim nx As Range
        Set nx = sr.NextStoryRange
        Do While Not nx Is Nothing
            leftAll = leftAll + nx.OMaths.Count
            Set nx = nx.NextStoryRange
        Loop
    Next sr

    ' the LAST-processed equation (first in the document) loses its OLE
    ' preview to a race with the final save: MathType flushes the object
    ' cache asynchronously, so drain the message queue before returning
    ' (observed as a blank first equation in short runs, 2026-11-03)
    Dim t3 As Single
    t3 = Timer
    Do While Timer - t3 < 1.5
        DoEvents
    Loop

    PyConvertAll = "ok=" & ok & " fail=" & fail & " left=" & leftAll & _
                   " shapes=" & doc.InlineShapes.Count & " stats=" & statHist
    Exit Function
eh:
    PyConvertAll = "VBAERR " & Err.Number & ": " & Err.Description & _
                   " (after ok=" & ok & " fail=" & fail & ")"
End Function

Private Function InsertMT(rng As Range, mml As String, ByRef statOut As Long) As Boolean
    Dim doc As Document
    Set doc = rng.Document
    InsertMT = False
    statOut = -12345

    On Error GoTo eh
    statOut = -2
    If rng.Start <> rng.End Then rng.Delete
    rng.Collapse wdCollapseStart
    statOut = -3
    rng.InsertFile BLANK_EQN, "MTBlankEqn", False, False, False
    rng.Expand wdWord

    statOut = -4
    Dim mtEqn As InlineShape
    Set mtEqn = Nothing
    If rng.InlineShapes.Count >= 1 Then
        Set mtEqn = rng.InlineShapes(1)
    End If
    If mtEqn Is Nothing Then statOut = -5: Exit Function

    statOut = -6
    mtEqn.OLEFormat.DoVerb 2
    Dim t As Single
    t = Timer
    Do While Timer - t < 0.4
        DoEvents
    Loop

    statOut = -7
    Dim obj As Object
    Set obj = mtEqn.OLEFormat.Object
    If obj Is Nothing Then statOut = -8888: Exit Function

    statOut = -8
    Dim b() As Byte
    b = mml
    statOut = -9
    Dim stat As Long
    stat = MTSetEqnFromLangStr(obj, 2, b(0), Len(mml))
    If stat <> 0 Then
        Dim t2 As Single
        t2 = Timer
        Do While Timer - t2 < 0.5
            DoEvents
        Loop
        stat = MTSetEqnFromLangStr(obj, 2, b(0), Len(mml))
    End If
    statOut = stat
    MTCloseOleObject 0, obj
    InsertMT = (stat = 0)
    Exit Function
eh:
    statOut = 9000 + Err.Number
End Function

' Non-ASCII chars die as '?' in the ANSI (mbcs) byte transfer to
' MTSetEqnFromLangStr. Numeric character refs are pure ASCII and the
' MathType MathML parser resolves them back to the real glyph.
Private Function EntityEncode(s As String) As String
    Dim i As Long, cp As Long, ch As String, out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        cp = AscW(ch)
        If cp < 0 Then cp = cp + 65536
        If cp > 127 Then
            out = out & "&#" & cp & ";"
        Else
            out = out & ch
        End If
    Next i
    EntityEncode = out
End Function
'''


class CrashDuringRun(RuntimeError):
    """Word process died mid-conversion (RPC failure). The output file holds
    a consistent checkpoint; the driver relaunches and resumes."""


class VbeUnavailable(RuntimeError):
    pass


CRASH_HRESULTS = {
    -2147023170,  # RPC_S_CALL_FAILED
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE
    -2147417848,  # CO_E_OBJNOTCONNECTED ("object disconnected")
    -2147418107,  # RPC_E_SYS_CALL_FAILED
}


def com_retry(fn, *args, **kw):
    last = None
    for _ in range(60):
        try:
            return fn(*args, **kw)
        except pythoncom.com_error as e:
            if e.hresult in (-2147418111, -2147417846):
                last = e
                time.sleep(0.5)
                continue
            raise
    raise last


def word_running():
    # tasklist output is GBK on zh-CN Windows; compare bytes, skip decoding
    out = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE"],
        capture_output=True).stdout
    return b"WINWORD.EXE" in out


def clear_word_locks():
    nlock = os.path.join(os.environ.get("APPDATA", ""),
                         "Microsoft", "Templates", "~$Normal.dotm")
    if os.path.isfile(nlock):
        try:
            os.remove(nlock)
        except OSError:
            pass


def stop_word(proc, word):
    try:
        com_retry(word.Quit)
    except Exception:
        pass
    # /embedding Word sometimes outlives Quit while our COM references
    # exist; a zombie keeps Normal.dotm locked and breaks the NEXT run's
    # COM attach, so make sure this process really dies
    for _ in range(10):
        if proc.poll() is not None:
            break
        time.sleep(1)
    if proc.poll() is None:
        proc.kill()
    time.sleep(1)
    clear_word_locks()


def launch_word(env):
    proc = subprocess.Popen([WINWORD, "/embedding"], env=env)
    word = None
    for _ in range(40):
        time.sleep(1)
        try:
            pythoncom.CoInitialize()
            word = wc.GetActiveObject("Word.Application")
            break
        except pythoncom.com_error:
            pass
    if word is None:
        proc.kill()
        raise SystemExit("cannot attach to WINWORD")
    word = wc.GetActiveObject("Word.Application")
    time.sleep(3)
    com_retry(setattr, word, "Visible", True)
    com_retry(setattr, word, "DisplayAlerts", 0)
    print("Word attached", flush=True)
    return proc, word


def build_addin(env):
    """One-time installation of the conversion macro as a Word Startup
    add-in. VBE automation is only needed HERE — after this succeeds, normal
    conversions never touch the VBA project system (its E_FAIL races were
    the flakiest part of the old pipeline)."""
    os.makedirs(STARTUP_DIR, exist_ok=True)
    for f in os.listdir(STARTUP_DIR):
        if f.startswith("PyMathTypeConv_") and f != ADDIN_NAME \
                and f.endswith(".dotm"):
            try:
                os.remove(os.path.join(STARTUP_DIR, f))
            except OSError:
                pass
    last = None
    for cycle in (1, 2):
        proc, word = launch_word(env)
        doc = None
        try:
            doc = com_retry(word.Documents.Add)
            vbproj = None
            for attempt in range(6):
                try:
                    time.sleep(2)
                    vbproj = doc.VBProject
                    break
                except pythoncom.com_error as e:
                    print("VBE busy (addin build), retry %d" % attempt,
                          flush=True)
                    if attempt == 5:
                        raise VbeUnavailable(str(e))
                    time.sleep(3)
            comp = vbproj.VBComponents.Add(1)
            comp.Name = "PyMathTypeConv"
            comp.CodeModule.AddFromString(VBA_CODE)
            try:
                vbproj.Name = "PyMathTypeConv"
            except pythoncom.com_error:
                pass  # name conflict — Run resolves unqualified names anyway
            com_retry(doc.SaveAs2, ADDIN_PATH, 15)  # wdFormatXMLTemplateMacroEnabled
            doc.Close(False)
            doc = None
            if os.path.isfile(ADDIN_PATH):
                print("宏插件已安装:", ADDIN_PATH, flush=True)
                return None
            last = "SaveAs2 未生成 dotm"
        except VbeUnavailable as e:
            last = str(e)
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            stop_word(proc, word)
        print("插件构建失败（第 %d 轮）: %s，清理锁后重试…" % (cycle, last),
              flush=True)
        time.sleep(3)
    return f"VBE 多轮重启后仍不可用: {last}"


def run_once(args, env, work, out, n_expected):
    """One Word session converting the working document. Raises
    CrashDuringRun if the process dies mid-conversion. A watchdog thread
    force-kills Word when the run exceeds a time budget scaled to the
    formula count — the first launched Word instance occasionally enters a
    ~20-min crawl (observed 2026-11-06); killing it converts the hang into
    a crash, which the driver auto-resumes from the checkpoint."""
    proc, word = launch_word(env)
    doc = None
    watchdog = None
    try:
        doc = com_retry(word.Documents.Open, work, ReadOnly=False,
                        AddToRecentFiles=False)
        com_retry(word.Activate)
        n_main = com_retry(lambda: doc.OMaths.Count)
        print("OMaths (main story):", n_main, flush=True)

        # watchdog: normal pace is ~0.8s/formula; budget = 3s/formula + 2min
        budget = max(240, n_main * 3 + 120) if n_main else 300
        watchdog = threading.Timer(
            budget, lambda: proc.kill(),
            kwargs=None)
        watchdog.daemon = True
        watchdog.start()

        t0 = time.time()
        try:
            result = com_retry(word.Run, "PyMathTypeConv.PyConvertAll",
                               args.limit)
        except pythoncom.com_error as e:
            if e.hresult in CRASH_HRESULTS or proc.poll() is not None:
                if watchdog:
                    watchdog.cancel()
                stop_word(proc, word)
                raise CrashDuringRun(
                    f"超时({budget}s)或进程死亡: {e}" if proc.poll() is not None
                    else str(e))
            raise
        if watchdog:
            watchdog.cancel()
        print("RESULT:", result, "(%.0fs)" % (time.time() - t0), flush=True)

        if result.startswith("ok="):
            time.sleep(1)  # let MathType finish flushing the last OLE cache
            com_retry(doc.Save)
            print("saved:", out, flush=True)
        else:
            # keep the last periodic-save checkpoint; changes after it are
            # discarded so the output stays consistent for a resume run
            print("进度保留在输出文件的最近检查点，重跑同一命令即可断点续转",
                  flush=True)
        doc.Close(False)
        doc = None
    finally:
        if watchdog:
            watchdog.cancel()
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        stop_word(proc, word)
    return result


# --- precomposed/combining accent decomposition -----------------------------
# MathType's MathML->MTEF translator ("WinAllBasicCodePages") lacks entries
# for precomposed accented chars (Ŷ U+0176 -> missing-glyph lozenge; Ã Â
# -> '?' via GBK-styled runs, Word-verified 2026-11-05) and renders
# combining-mark sequences as a separate caret. Word's own m:acc accent
# structure renders all of these perfectly, so rewrite every affected m:r
# run into m:acc BEFORE handing the document to the engine. NFD-first so
# precomposed chars and raw combining sequences both work.

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
M = "{%s}" % M_NS

COMBINING_ACCENT = {
    "\u0300": "`", "\u0301": "\u00b4", "\u0302": "^", "\u0303": "~",
    "\u0304": "\u00af", "\u0306": "\u02d8", "\u0307": "\u02d9",
    "\u0308": "\u00a8", "\u030a": "\u02da", "\u030c": "\u02c7",
    "\u20d7": "\u2192",
}


def _accent_segments(text):
    """[('text', s) | ('acc', base, mark)] segments covering the whole text,
    or None when nothing is decomposable. Bails out (None) if the NFD form
    contains combining marks outside our table — leaking those into plain
    runs would render worse than the precomposed original."""
    src = unicodedata.normalize("NFD", text)
    if any(unicodedata.combining(ch) and ch not in COMBINING_ACCENT
           for ch in src):
        return None
    segs = []
    buf = []
    for ch in src:
        if ch in COMBINING_ACCENT and buf:
            base = buf.pop()
            if buf:
                segs.append(("text", "".join(buf)))
                buf = []
            segs.append(("acc", base, COMBINING_ACCENT[ch]))
        else:
            buf.append(ch)
    if buf:
        segs.append(("text", "".join(buf)))
    return segs if any(s[0] == "acc" for s in segs) else None


def _mk_run(rpr, text):
    r = etree.Element(M + "r")
    if rpr is not None:
        r.append(etree.fromstring(etree.tostring(rpr)))
    t = etree.SubElement(r, M + "t")
    t.text = text
    return r


def _mk_acc(rpr, base, mark):
    acc = etree.Element(M + "acc")
    pr = etree.SubElement(acc, M + "accPr")
    chr_el = etree.SubElement(pr, M + "chr")
    chr_el.set(M + "val", mark)
    e = etree.SubElement(acc, M + "e")
    e.append(_mk_run(rpr, base))
    return acc


def decompose_accents(path):
    """Rewrite m:r runs holding decomposable accent chars into m:acc
    structures, across every story part. Idempotent."""
    zin = zipfile.ZipFile(path)
    items = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    changed = 0
    for name in list(items):
        if not (name == "word/document.xml" or re.match(
                r"word/(footnotes|endnotes|comments|header\d*|footer\d*)\.xml$",
                name)):
            continue
        try:
            root = etree.fromstring(items[name])
        except etree.XMLSyntaxError:
            continue
        touched = False
        for t_el in root.iter(M + "t"):
            text = t_el.text or ""
            segs = _accent_segments(text)
            if segs is None:
                continue
            r_el = t_el.getparent()
            if r_el is None or r_el.tag != M + "r":
                continue
            parent = r_el.getparent()
            if parent is None:
                continue
            rpr = r_el.find(M + "rPr")
            pos = list(parent).index(r_el)
            new_els = []
            for seg in segs:
                if seg[0] == "text":
                    if seg[1]:
                        new_els.append(_mk_run(rpr, seg[1]))
                else:
                    new_els.append(_mk_acc(rpr, seg[1], seg[2]))
            parent.remove(r_el)
            for k, el in enumerate(new_els):
                parent.insert(pos + k, el)
            changed += 1
            touched = True
        if touched:
            items[name] = etree.tostring(
                root, xml_declaration=True, standalone=True, encoding="UTF-8")
    if changed:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zo:
            for n, blob in items.items():
                zo.writestr(n, blob)
    return changed


def prepare_file(inp, out):
    """Copy input -> output (the working copy), apply accent decomposition.
    Returns True when the file needs conversion, False when already done."""
    resume = False
    if os.path.isfile(out):
        in_ds, in_omath = census(inp)
        out_ds, out_left = census(out)
        if out_left == 0 and out_ds == in_ds + in_omath and (in_ds + in_omath) > 0:
            print(f"跳过 {os.path.basename(inp)}: 输出已完成"
                  f"（{out_ds} 个 MathType 对象）", flush=True)
            return False
        if out_left > 0 and out_left + (out_ds - in_ds) == in_omath:
            resume = True
            print(f"断点续转: 输出已含 {out_ds} 个 MathType 对象、"
                  f"剩 {out_left} 个原生公式，从中继续", flush=True)
        else:
            raise SystemExit(
                f"输出与输入不匹配（已转 {out_ds} + 剩余 {out_left} "
                f"!= 输入待转 {in_omath}），请删除输出或换 -o 路径。")
    if not resume:
        # checkpoint base + accent decomposition happen on the OUTPUT copy;
        # the input file is never modified
        shutil.copyfile(inp, out)
    n_dec = decompose_accents(out)
    if n_dec:
        print(f"重音分解: {n_dec} 个含预组合/组合字符的 run 重写为 m:acc",
              flush=True)
    return True


def main():
    ap = argparse.ArgumentParser(
        description="Convert native OMML equations (all stories) to MathType OLE objects.")
    ap.add_argument("input", nargs="+",
                    help=".docx file(s) whose equations get converted "
                    "(inputs are left unchanged; results go to outputs)")
    ap.add_argument("-o", "--output", default=None,
                    help="output .docx (single input only; default "
                    "<input>_MathType.docx)")
    ap.add_argument("--kill-word", action="store_true",
                    help="force-kill every running WINWORD.EXE first "
                    "(unsaved work in other Word documents is lost)")
    ap.add_argument("--limit", type=int, default=0,
                    help="convert only the first N formulas per file "
                    "(smoke test; 0=all)")
    ap.add_argument("--rebuild-addin", action="store_true",
                    help="force-rebuild the macro add-in from the current "
                    "VBA source, then convert")
    args = ap.parse_args()

    if len(args.input) > 1 and args.output:
        raise SystemExit("-o 只支持单个输入文件")
    files = []
    for raw in args.input:
        inp = os.path.abspath(raw)
        if not os.path.isfile(inp):
            raise SystemExit(f"input not found: {inp}")
        out = os.path.abspath(args.output) if args.output and len(args.input) == 1 \
            else os.path.splitext(inp)[0] + "_MathType.docx"
        if os.path.normcase(out) == os.path.normcase(inp):
            raise SystemExit("output must differ from input")
        files.append((inp, out))

    if word_running():
        if args.kill_word:
            subprocess.run(["taskkill", "/IM", "WINWORD.EXE", "/F"],
                           capture_output=True)
            time.sleep(2)
        else:
            raise SystemExit(
                "WINWORD.EXE is already running (leftover from a crashed run, "
                "or your own session). Close it and retry, or pass --kill-word "
                "to force-kill all Word processes.")
    for inp, out in files:
        for base in (inp, out):
            lock = os.path.join(os.path.dirname(base),
                                "~$" + os.path.basename(base))
            if os.path.isfile(lock):
                try:
                    os.remove(lock)
                except OSError:
                    pass
    clear_word_locks()

    env = dict(os.environ)
    env["PATH"] = os.pathsep.join(
        [MT + r"\MathPage\64", MT + r"\System\64", env.get("PATH", "")])

    if args.rebuild_addin and os.path.isfile(ADDIN_PATH):
        os.remove(ADDIN_PATH)
    if not os.path.isfile(ADDIN_PATH):
        print("首次运行：安装宏插件（一次性，之后无需 VBE 访问）…", flush=True)
        err = build_addin(env)
        if err:
            raise SystemExit("宏插件安装失败: " + err)

    for inp, out in files:
        print(f"== {os.path.basename(inp)}", flush=True)
        if not prepare_file(inp, out):
            continue
        attempts = 0
        while True:
            try:
                in_ds, in_om = census(inp)
                run_once(args, env, out, out, in_ds + in_om)
                break
            except CrashDuringRun as e:
                attempts += 1
                print(f"Word 进程崩溃（第 {attempts} 次）: {e}", flush=True)
                if attempts >= 3:
                    raise SystemExit("Word 连续崩溃 3 次，中止。"
                                     "重跑同一命令可从最近检查点续转。")
                out_ds, out_left = census(out)
                in_ds, in_om = census(inp)
                if out_left == 0 and out_ds == in_ds + in_om:
                    print("崩溃时已全部转换完成", flush=True)
                    break
                if not (out_left + (out_ds - in_ds) == in_om):
                    raise SystemExit("崩溃后输出状态不一致，请检查输出文件。")
                print("自动从检查点续转…", flush=True)


if __name__ == "__main__":
    main()
