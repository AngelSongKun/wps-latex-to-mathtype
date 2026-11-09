# -*- coding: utf-8 -*-
"""One-time environment fixes that made MathType's OMML->MathType conversion
work on Windows 11 (verified 2026-10-24 on a real 228-equation thesis).

Idempotent - safe to run multiple times. All changes are per-user (HKCU),
no admin rights required.

Fixes:
1. MSXML2.DOMDocument30 ProgID    (Win11 24H2+ dropped versioned MSXML ProgIDs)
2. MSXML 3.0 TypeLib registration (needed for early binding in MathType's VBA)
3. Word AccessVBOM = 1            (trust VBA project object model, lets the
                                    converter inject a module into Normal)
4. User PATH += MathType MathPage/System dirs (Declare Lib "MathPage.WLL")
"""
import os
import winreg

MT = r"C:\Program Files (x86)\MathType"
PROGID_CLSID = "{F5078F32-C551-11D3-89B9-0000F81FE221}"  # DOMDocument 3.0


def main():
    # 1) ProgID -> CLSID
    k = winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                         r"Software\Classes\MSXML2.DOMDocument30\CLSID")
    winreg.SetValueEx(k, "", 0, winreg.REG_SZ, PROGID_CLSID)
    winreg.CloseKey(k)
    print("OK  HKCU ProgID MSXML2.DOMDocument30 ->", PROGID_CLSID)

    # 2) TypeLib for early binding
    import pythoncom
    tlb = pythoncom.LoadTypeLib(os.path.join(os.environ["WINDIR"], "System32", "msxml3.dll"))
    la = tlb.GetLibAttr()
    guid = str(la[0]).strip("{}")
    base = "Software\\Classes\\TypeLib\\{%s}\\%d.0" % (guid, la[3])
    for sub, val, typ in [
        ("\\0\\win64", os.path.join(os.environ["WINDIR"], "System32", "msxml3.dll"), winreg.REG_SZ),
        ("\\0\\win32", os.path.join(os.environ["WINDIR"], "SysWOW64", "msxml3.dll"), winreg.REG_SZ),
        ("\\FLAGS", 0, winreg.REG_DWORD),
    ]:
        k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, base + sub)
        winreg.SetValueEx(k, "", 0, typ, val)
        winreg.CloseKey(k)
    print("OK  HKCU TypeLib {%s} %d.0" % (guid, la[3]))

    # 3) Word: trust VBA project object model
    for ver in ("16.0",):
        k = winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Office\%s\Word\Security" % ver)
        winreg.SetValueEx(k, "AccessVBOM", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(k)
    print("OK  AccessVBOM = 1")

    # 4) user PATH += MathType dirs
    env = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_ALL_ACCESS)
    try:
        old, _ = winreg.QueryValueEx(env, "Path")
    except FileNotFoundError:
        old = ""
    add = [MT + r"\MathPage\64", MT + r"\System\64"]
    parts = old.split(";") if old else []
    changed = False
    for d in add:
        if d.lower() not in [p.lower() for p in parts]:
            parts.append(d)
            changed = True
    if changed:
        winreg.SetValueEx(env, "Path", 0, winreg.REG_EXPAND_SZ, ";".join(parts))
        print("OK  user PATH +=", add)
    else:
        print("OK  PATH already contains MathType dirs")
    winreg.CloseKey(env)

    import ctypes
    ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x1A, 0, "Environment", 2, 5000,
                                             ctypes.byref(ctypes.c_ulong()))
    print("OK  environment broadcast sent")


if __name__ == "__main__":
    main()
