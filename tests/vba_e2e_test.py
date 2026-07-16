# -*- coding: utf-8 -*-
# End-to-end test of the VBA SharePoint Duplicate Checker via Excel COM.
# Scenarios:
#   T1 normal compare: New / Missing / Duplicate, key columns named
#      differently between Master ("ID") and Current ("รหัส") to prove
#      column-name independence
#   T2 unknown key column -> must hard-error with a readable message
import os, re, shutil, sys, tempfile, traceback
import win32com.client
from openpyxl import Workbook

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VBA_DIR = os.path.join(REPO, "vba")
SCRATCH = os.path.join(tempfile.gettempdir(), "sharepoint_dup_checker_e2e")
VB_OBJ_ERR = -2147221504  # vbObjectError


def prep_dirs():
    if os.path.exists(SCRATCH):
        shutil.rmtree(SCRATCH)
    os.makedirs(os.path.join(SCRATCH, "files"))
    os.makedirs(os.path.join(SCRATCH, "mods"))


def build_sample_files():
    master_path = os.path.join(SCRATCH, "files", "Master.xlsx")
    current_path = os.path.join(SCRATCH, "files", "Current.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["ID", "Name"])
    for row in [("M1", "Alice"), ("M2", "Bob"), ("M3", "Carol"), ("M4", "Dave")]:
        ws.append(row)
    wb.save(master_path)

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["รหัส", "Name"])  # "รหัส" — deliberately different header from Master's "ID"
    for row in [("M2", "Bob"), ("M3", "Carol"), ("M4", "Dave"),
                ("M4", "Dave"), ("N1", "Eve")]:
        ws.append(row)
    wb.save(current_path)

    return master_path, current_path


def patch_modules():
    """Patch MsgBox -> cell logging so macros run unattended; write as cp874
    (Thai ANSI) because VBE imports .bas as ANSI."""
    mods = []
    for name in ("modUtils", "modConfig", "modCompare"):
        src = open(os.path.join(VBA_DIR, name + ".bas"), encoding="utf-8").read()
        src = re.sub(
            r'MsgBox "เปรียบเทียบเสร็จแล้ว(?:.|\n)*?vbInformation, "SharePoint Duplicate Checker"',
            'ThisWorkbook.Worksheets("Config").Range("Z1").Value = '
            '"OK New=" & newCount & " Missing=" & missingCount & " Dup=" & dupCount',
            src)
        src = re.sub(
            r'MsgBox "เกิดข้อผิดพลาดใน(?:.|\n)*?vbCritical, "SharePoint Duplicate Checker"',
            'ThisWorkbook.Worksheets("Config").Range("Z2").Value = '
            '"ERR|" & errNumber & "|" & PROC_NAME & "|" & errDescription',
            src)
        out = os.path.join(SCRATCH, "mods", name + ".bas")
        open(out, "w", encoding="cp874", errors="replace", newline="\r\n").write(src)
        mods.append(out)
    return mods


def build_workbook(xl, mods, master_path, current_path):
    wb = xl.Workbooks.Add()
    while wb.Worksheets.Count > 1:
        wb.Worksheets(wb.Worksheets.Count).Delete()
    ws = wb.Worksheets(1)
    ws.Name = "Config"
    ws.Range("A1").Value = "Key"
    ws.Range("B1").Value = "Value"
    rows = [
        ("MasterFilePath", master_path),
        ("CurrentFilePath", current_path),
        ("SheetName", "Sheet1"),
        ("MasterKeyColumn", "ID"),
        ("CurrentKeyColumn", "รหัส"),
    ]
    for i, (k, v) in enumerate(rows, start=2):
        ws.Range(f"A{i}").Value = k
        ws.Range(f"B{i}").Value = v
    lo = ws.ListObjects.Add(1, ws.Range(f"A1:B{len(rows) + 1}"), None, 1)
    lo.Name = "tblConfig"

    for m in mods:
        wb.VBProject.VBComponents.Import(m)
    return wb


def set_config(wb, key, value):
    lo = wb.Worksheets("Config").ListObjects("tblConfig")
    for r in range(1, lo.ListRows.Count + 1):
        if str(lo.DataBodyRange.Cells(r, 1).Value).strip() == key:
            lo.DataBodyRange.Cells(r, 2).Value = value
            return
    raise RuntimeError("setting not found: " + key)


def clear_log(wb):
    ws = wb.Worksheets("Config")
    ws.Range("Z1").Value = None
    ws.Range("Z2").Value = None


def read_log(wb):
    ws = wb.Worksheets("Config")
    return ws.Range("Z1").Value, ws.Range("Z2").Value


def read_report(wb):
    for sh in wb.Worksheets:
        for lo in sh.ListObjects:
            if lo.Name == "ReportTable":
                if lo.DataBodyRange is None:
                    return set()
                n = lo.DataBodyRange.Rows.Count
                return {(lo.DataBodyRange.Cells(r, 1).Value,
                         lo.DataBodyRange.Cells(r, 2).Value) for r in range(1, n + 1)}
    return None


results = []


def check(label, cond, detail=""):
    results.append((label, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + label + ("  | " + str(detail) if detail else ""))


def main():
    prep_dirs()
    master_path, current_path = build_sample_files()
    mods = patch_modules()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        try:
            wb = build_workbook(xl, mods, master_path, current_path)
        except Exception:
            print("VBProject import blocked -> enabling AccessVBOM and retrying")
            xl.Quit()
            import winreg
            ver = "16.0"
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                rf"Software\Microsoft\Office\{ver}\Excel\Security")
            winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
            winreg.CloseKey(key)
            xl = win32com.client.DispatchEx("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
            wb = build_workbook(xl, mods, master_path, current_path)

        # ---- T1: normal compare, key columns named differently ----
        clear_log(wb)
        xl.Run("RunCompareFiles")
        z1, z2 = read_log(wb)
        report = read_report(wb)
        check("T1 ran without error", z2 is None, z2)
        check("T1 success log", z1 == "OK New=1 Missing=1 Dup=1", z1)
        expected = {("M4", "Duplicate (2x)"), ("N1", "New"), ("M1", "Missing")}
        check("T1 report matches expected New/Missing/Duplicate", report == expected, report)

        # ---- T2: unknown key column must hard-error ----
        set_config(wb, "CurrentKeyColumn", "NoSuchColumn")
        clear_log(wb)
        xl.Run("RunCompareFiles")
        z1, z2 = read_log(wb)
        got = str(z2).split("|") if z2 else []
        expected_err = VB_OBJ_ERR + 612
        check("T2 unknown key column rejected",
              len(got) >= 2 and got[0] == "ERR" and int(got[1]) == expected_err, z2)

        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()

    fails = [r for r in results if not r[1]]
    print(f"\n===== {len(results) - len(fails)}/{len(results)} passed =====")
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        sys.exit(2)
