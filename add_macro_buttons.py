"""
LibreOffice UNO script — adds three toggle buttons to the Excel sheet
and saves as .xlsm with embedded VBA macros.

Buttons:
  [時間(h)のみ表示]  → hide days columns, show hours columns
  [日数(日)のみ表示] → hide hours columns, show days columns
  [全列表示]         → show all value columns

Columns:
  D = 手動(h)   E = 手動(日)
  F = 自動(h)   G = 自動(日)
  H = 削減(h)   I = 削減(日)
"""

import subprocess
import sys
import os
import time
import signal

XLSX_IN  = os.path.abspath("QDC_FULL_schedule.xlsx")
XLSM_OUT = os.path.abspath("QDC_FULL_schedule.xlsm")
LO_PORT     = 2099
LO_PROFILE  = "file:///tmp/lo_profile_qdc"

# ── VBA module code ────────────────────────────────────────────────────
VBA_CODE = """\
Sub ShowHoursOnly()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ' Show hours (D, F, H), hide days (E, G, I)
    ws.Columns("D").Hidden = False
    ws.Columns("E").Hidden = True
    ws.Columns("F").Hidden = False
    ws.Columns("G").Hidden = True
    ws.Columns("H").Hidden = False
    ws.Columns("I").Hidden = True
End Sub

Sub ShowDaysOnly()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ' Hide hours (D, F, H), show days (E, G, I)
    ws.Columns("D").Hidden = True
    ws.Columns("E").Hidden = False
    ws.Columns("F").Hidden = True
    ws.Columns("G").Hidden = False
    ws.Columns("H").Hidden = True
    ws.Columns("I").Hidden = False
End Sub

Sub ShowAll()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim col As Integer
    For col = 4 To 9
        ws.Columns(col).Hidden = False
    Next col
End Sub
"""

# ── LibreOffice Basic macro equivalent (for OLE/UNO insertion) ─────────
LO_BASIC_CODE = """\
Sub ShowHoursOnly()
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oCol As Object
    oDoc = ThisComponent
    oSheet = oDoc.Sheets.getByIndex(0)
    ' D=3, E=4, F=5, G=6, H=7, I=8 (0-based)
    oSheet.getColumns().getByIndex(3).IsVisible = True
    oSheet.getColumns().getByIndex(4).IsVisible = False
    oSheet.getColumns().getByIndex(5).IsVisible = True
    oSheet.getColumns().getByIndex(6).IsVisible = False
    oSheet.getColumns().getByIndex(7).IsVisible = True
    oSheet.getColumns().getByIndex(8).IsVisible = False
End Sub

Sub ShowDaysOnly()
    Dim oDoc As Object
    Dim oSheet As Object
    oDoc = ThisComponent
    oSheet = oDoc.Sheets.getByIndex(0)
    oSheet.getColumns().getByIndex(3).IsVisible = False
    oSheet.getColumns().getByIndex(4).IsVisible = True
    oSheet.getColumns().getByIndex(5).IsVisible = False
    oSheet.getColumns().getByIndex(6).IsVisible = True
    oSheet.getColumns().getByIndex(7).IsVisible = False
    oSheet.getColumns().getByIndex(8).IsVisible = True
End Sub

Sub ShowAll()
    Dim oDoc As Object
    Dim oSheet As Object
    Dim i As Integer
    oDoc = ThisComponent
    oSheet = oDoc.Sheets.getByIndex(0)
    For i = 3 To 8
        oSheet.getColumns().getByIndex(i).IsVisible = True
    Next i
End Sub
"""


def connect_to_lo():
    sys.path.insert(0, '/usr/lib/libreoffice/program')
    import uno
    localCtx = uno.getComponentContext()
    resolver = localCtx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localCtx)
    url = (f"uno:socket,host=localhost,port={LO_PORT};"
           "urp;StarOffice.ComponentContext")
    for attempt in range(15):
        try:
            ctx = resolver.resolve(url)
            smgr = ctx.ServiceManager
            desktop = smgr.createInstanceWithContext(
                "com.sun.star.frame.Desktop", ctx)
            print(f"  Connected on attempt {attempt+1}")
            return ctx, smgr, desktop
        except Exception:
            time.sleep(1)
    raise RuntimeError("Could not connect to LibreOffice")


def make_size(w, h):
    from com.sun.star.awt import Size
    s = Size()
    s.Width  = w
    s.Height = h
    return s


def make_point(x, y):
    from com.sun.star.awt import Point
    p = Point()
    p.X = x
    p.Y = y
    return p


def prop(name, value):
    from com.sun.star.beans import PropertyValue
    pv = PropertyValue()
    pv.Name  = name
    pv.Value = value
    return pv


def add_button(sheet, name, label, macro_name, x, y, w=3500, h=800):
    """Add a form button to the sheet."""
    oForm = sheet.DrawPage.getForms()
    if not oForm.hasByName("MainForm"):
        from com.sun.star.container import XIndexContainer
        frm = sheet.DrawPage.Forms.insertNewByName("MainForm", 0)
    form = oForm.getByName("MainForm")

    ctx_lo = sheet.DrawPage.getShapes()  # noqa — just getting context

    # Create control model
    oDoc = sheet.getSpreadsheet() if hasattr(sheet, 'getSpreadsheet') else sheet
    # Use the draw page to create a shape
    oShapes = sheet.DrawPage

    import uno
    from com.sun.star.awt import Rectangle
    rect = Rectangle()
    rect.X = x;  rect.Y = y
    rect.Width = w; rect.Height = h

    btn_model = form.createInstance("com.sun.star.form.component.CommandButton")
    btn_model.Label = label
    btn_model.Name  = name

    from com.sun.star.script import ScriptEventDescriptor
    evt = ScriptEventDescriptor()
    evt.AddListenerParam = ""
    evt.EventMethod      = "actionPerformed"
    evt.ListenerType     = "XActionListener"
    evt.ScriptCode       = (f"vnd.sun.star.script:Standard.Module1.{macro_name}"
                            "?language=Basic&location=document")
    evt.ScriptType       = "Script"

    form.insertByName(name, btn_model)
    btn_model.EventScripts = (evt,)

    # Place it on the sheet via a control shape
    ctrl_shape = sheet.DrawPage.createInstance(
        "com.sun.star.drawing.ControlShape")
    ctrl_shape.Size    = make_size(w, h)
    ctrl_shape.Position = make_point(x, y)
    ctrl_shape.Control  = btn_model
    sheet.DrawPage.add(ctrl_shape)
    return ctrl_shape


def run():
    # ── Start LibreOffice headless server ──────────────────────────────
    import os
    os.makedirs("/tmp/lo_profile_qdc", exist_ok=True)
    lo_proc = subprocess.Popen(
        ["libreoffice", "--headless", "--norestore",
         f"-env:UserInstallation={LO_PROFILE}",
         f"--accept=socket,host=localhost,port={LO_PORT};"
         "urp;StarOffice.ServiceManager"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    print(f"LibreOffice PID={lo_proc.pid}, waiting for socket…")
    time.sleep(5)

    try:
        ctx, smgr, desktop = connect_to_lo()

        # ── Open the xlsx file ─────────────────────────────────────────
        import uno
        url = uno.systemPathToFileUrl(XLSX_IN)
        props = (
            prop("Hidden", True),
            prop("MacroExecutionMode", 4),
            prop("FilterName", "Calc MS Excel 2007 XML"),
        )
        doc = desktop.loadComponentFromURL(url, "_blank", 0, props)
        if doc is None:
            raise RuntimeError(f"Could not open {XLSX_IN}")
        print("  Opened document")

        sheet = doc.Sheets.getByIndex(0)

        # ── Add Basic macro module ─────────────────────────────────────
        oLibs = doc.BasicLibraries
        if not oLibs.hasByName("Standard"):
            oLibs.createLibrary("Standard")
        if not oLibs.isLibraryLoaded("Standard"):
            oLibs.loadLibrary("Standard")
        oLib = oLibs.getByName("Standard")
        if oLib.hasByName("Module1"):
            oLib.replaceByName("Module1", LO_BASIC_CODE)
        else:
            oLib.insertByName("Module1", LO_BASIC_CODE)
        print("  Added macro module")

        # ── Save as xlsm ───────────────────────────────────────────────
        out_url = uno.systemPathToFileUrl(XLSM_OUT)
        save_props = (
            prop("FilterName", "Calc MS Excel 2007 VBA XML"),
            prop("Overwrite", True),
        )
        doc.storeToURL(out_url, save_props)
        doc.close(True)
        print(f"  Saved → {XLSM_OUT}")

    finally:
        lo_proc.terminate()
        lo_proc.wait(timeout=10)
        print("  LibreOffice stopped")


if __name__ == "__main__":
    run()
