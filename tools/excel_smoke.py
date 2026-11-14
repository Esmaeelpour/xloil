"""
Opens Excel, loads the installed xlOil add-in and checks that Python answers worksheet formulas with exact values.

Used by the install smoke test (.github/workflows/install-smoke.yml) on a clean machine, and by hand:

    python tools/excel_smoke.py            # loads %APPDATA%\\Microsoft\\Excel\\XLSTART\\xlOil.xll

Excel started through COM does not open XLSTART add-ins, so the add-in is loaded with RegisterXLL. A watchdog
ends the run (exit 2) if Excel hangs, e.g. on a first-run or activation dialog. Exit 0: every check passed.
"""
import os
import sys
import threading
import time
from pathlib import Path

TIMEOUT_S = 300


def _watchdog():
    time.sleep(TIMEOUT_S)
    print(f"FAIL Excel did not finish within {TIMEOUT_S}s (a dialog may be open)", flush=True)
    os._exit(2)


def main(xll: Path) -> int:
    import pythoncom
    import win32com.client

    threading.Thread(target=_watchdog, daemon=True).start()
    pythoncom.CoInitialize()
    print(f"add-in: {xll} ({'present' if xll.exists() else 'MISSING'})", flush=True)
    if not xll.exists():
        return 1
    app = win32com.client.DispatchEx("Excel.Application")
    app.DisplayAlerts = False
    print(f"Excel {app.Version} started", flush=True)
    failed = 0
    try:
        loaded = app.RegisterXLL(str(xll))
        print(f"{'OK  ' if loaded else 'FAIL'} RegisterXLL -> {loaded}", flush=True)
        failed += not loaded
        wb = app.Workbooks.Add()
        ws = wb.Worksheets(1)
        cases = [
            ("A1", '=INDEX(xloVersion(),1,1)', "Version"),
            ("A2", '=xloAttr("python works","upper")', "PYTHON WORKS"),
            ("A3", '=xloAttr("a,b,c","count",,",")', 2.0),
            ("A4", '=xloAttr(xloAttrObj("x","__class__"),"__module__")', "builtins"),
        ]
        values = {}
        for _ in range(30):                 # Python starts in the background: wait until the formulas answer
            for cell, formula, _want in cases:
                ws.Range(cell).Formula = formula
            app.Calculate()
            values = {cell: ws.Range(cell).Value for cell, _f, _w in cases}
            if all(values[c] == w for c, _f, w in cases):
                break
            time.sleep(2)
        for cell, formula, want in cases:
            ok = values.get(cell) == want
            failed += not ok
            print(f"{'OK  ' if ok else 'FAIL'} {formula} -> {values.get(cell)!r} (want {want!r})", flush=True)
        wb.Close(SaveChanges=False)
    finally:
        app.Quit()
    print("RESULT:", "all checks passed" if not failed else f"{failed} check(s) failed", flush=True)
    return 1 if failed else 0


if __name__ == "__main__":
    default = Path(os.environ["APPDATA"]) / "Microsoft" / "Excel" / "XLSTART" / "xlOil.xll"
    sys.exit(main(Path(sys.argv[1]) if len(sys.argv) > 1 else default))
