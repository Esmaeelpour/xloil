"""
    Builds xlOil.xlam: an .xlam ribbon for xlOil, for computers whose Office policy
    blocks unsigned COM add-ins (which the default `xloil.xloil_ribbon` tab uses).

    Every control calls a small VBA dispatcher which forwards to the xlOil function
    `xloRibbonCall` (xloil/xlam_ribbon.py). VBA is needed because Excel resolves a
    workbook ribbon's callbacks through that workbook's VBA project.

    Built entirely in Python (tools/vba_project.py writes the VBA part), so it
    needs neither Excel nor "Trust access to the VBA project object model".
    Run from the repo root:

        python tools/build_xlam.py

    Writes libs/xlOil_Python/Package/xloil/data/xlOil.xlam.
"""
import os
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import quoteattr

ROOT = Path(__file__).resolve().parent.parent
PACKAGE = ROOT / "libs" / "xlOil_Python" / "Package"
OUT = PACKAGE / "xloil" / "data" / "xlOil.xlam"

VBA = r'''' xlOil's settings ribbon. Each control forwards to the xlOil function
' xloRibbonCall(kind, tag, ...), implemented in Python (xloil/xlam_ribbon.py).
Option Explicit

Private gRibbon As Object

Private Function Xlo(kind As String, control As Object, Optional a As Variant, Optional b As Variant) As Variant
    On Error GoTo NotLoaded
    If IsMissing(a) Then
        Xlo = Application.Run("xloRibbonCall", kind, control.Tag)
    ElseIf IsMissing(b) Then
        Xlo = Application.Run("xloRibbonCall", kind, control.Tag, a)
    Else
        Xlo = Application.Run("xloRibbonCall", kind, control.Tag, a, b)
    End If
    If IsError(Xlo) Then Xlo = Empty
    Exit Function
NotLoaded:
    Xlo = Empty
End Function

Private Sub Refresh()
    If Not gRibbon Is Nothing Then gRibbon.Invalidate
End Sub

Private Sub Act(kind As String, control As Object, Optional a As Variant, Optional b As Variant)
    On Error GoTo NotLoaded
    If IsMissing(a) Then
        Application.Run "xloRibbonCall", kind, control.Tag
    ElseIf IsMissing(b) Then
        Application.Run "xloRibbonCall", kind, control.Tag, a
    Else
        Application.Run "xloRibbonCall", kind, control.Tag, a, b
    End If
    Refresh
    Exit Sub
NotLoaded:
    MsgBox "xlOil isn't loaded (or is still starting up). Check that xlOil.xll is in your XLSTART folder " & _
           "and that 'xloil.xlam_ribbon' is in LoadModules, then restart Excel." & vbCrLf & vbCrLf & _
           "Details: " & Err.Description, vbExclamation, "xlOil"
End Sub

Public Sub xlo_OnLoad(ribbon As Object)
    Set gRibbon = ribbon
End Sub

Public Sub xlo_OnAction(control As Object)
    Act "action", control
End Sub

Public Sub xlo_OnToggle(control As Object, pressed As Boolean)
    Act "toggle", control, pressed
End Sub

Public Sub xlo_OnChange(control As Object, text As String)
    Act "change", control, text
End Sub

Public Sub xlo_OnSelect(control As Object, id As String, index As Integer)
    Act "select", control, id, index
End Sub

Public Sub xlo_GetText(control As Object, ByRef returnedVal)
    returnedVal = CStr(Xlo("text", control))
End Sub

Public Sub xlo_GetPressed(control As Object, ByRef returnedVal)
    Dim v As Variant: v = Xlo("pressed", control)
    returnedVal = (Not IsEmpty(v)) And CBool(v)
End Sub

Public Sub xlo_GetEnabled(control As Object, ByRef returnedVal)
    Dim v As Variant: v = Xlo("enabled", control)
    returnedVal = IsEmpty(v) Or CBool(v)
End Sub

Public Sub xlo_GetVisible(control As Object, ByRef returnedVal)
    Dim v As Variant: v = Xlo("visible", control)
    returnedVal = (Not IsEmpty(v)) And CBool(v)
End Sub

Public Sub xlo_GetSupertip(control As Object, ByRef returnedVal)
    returnedVal = CStr(Xlo("supertip", control))
End Sub

Public Sub xlo_GetItemCount(control As Object, ByRef returnedVal)
    Dim v As Variant: v = Xlo("count", control)
    If IsEmpty(v) Then returnedVal = 0 Else returnedVal = CLng(v)
End Sub

Public Sub xlo_GetItemLabel(control As Object, index As Integer, ByRef returnedVal)
    returnedVal = CStr(Xlo("itemlabel", control, index))
End Sub

Public Sub xlo_GetSelectedItemIndex(control As Object, ByRef returnedVal)
    Dim v As Variant: v = Xlo("selected", control)
    If IsEmpty(v) Then returnedVal = 0 Else returnedVal = CLng(v)
End Sub
'''


def ribbon_xml() -> str:
    sys.path.insert(0, str(PACKAGE))
    os.environ.setdefault("XLOIL_STUBS", "1")
    from xloil.features import FEATURES

    def tip(name, value):
        return f" {name}={quoteattr(value)}" if value else ""

    def edit(tag, label, size, screentip, supertip=""):
        return (f'<editBox id="eb_{tag}" tag="{tag}" label={quoteattr(label)} sizeString={quoteattr(size)}'
                f'{tip("screentip", screentip)}{tip("supertip", supertip)} '
                f'getText="xlo_GetText" onChange="xlo_OnChange"/>')

    def drop(tag, label, screentip, supertip=""):
        return (f'<dropDown id="dd_{tag}" tag="{tag}" label={quoteattr(label)}'
                f'{tip("screentip", screentip)}{tip("supertip", supertip)} getItemCount="xlo_GetItemCount" getItemLabel="xlo_GetItemLabel" '
                f'getSelectedItemIndex="xlo_GetSelectedItemIndex" onAction="xlo_OnSelect"/>')

    def button(tag, label, image, screentip, supertip="", size="large"):
        return (f'<button id="bt_{tag}" tag="{tag}" size="{size}" label={quoteattr(label)} imageMso="{image}"'
                f'{tip("screentip", screentip)}{tip("supertip", supertip)} onAction="xlo_OnAction"/>')

    features = "".join(
        f'<checkBox id="cb_{f.key}" tag="feature:{f.key}" label={quoteattr(f.label)} '
        f'getPressed="xlo_GetPressed" getEnabled="xlo_GetEnabled" getSupertip="xlo_GetSupertip" '
        f'onAction="xlo_OnToggle"/>' for f in FEATURES)

    return f'''<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" onLoad="xlo_OnLoad">
  <ribbon><tabs><tab id="xloPyTab" label="xlOil Py">
    <group id="g_env" label="Environment">
      {drop("env", "Environment", "Python environments found on this computer", "Changes take effect after restarting Excel")}
      {edit("pyexe", "PYTHONEXECUTABLE", "c:/a/path/is/this/size", "The python.exe xlOil uses", "Changes take effect after restarting Excel")}
      {edit("pypath", "PYTHONPATH", "c:/a/path/is/this/size", "Semi-colon separated module search folders", "Prefer Search Paths for your own folders")}
    </group>
    <group id="g_mod" label="Modules">
      <menu id="mn_features" label="Features" size="large" imageMso="AddInManager" screentip="Turn optional xlOil features on or off" supertip="Adds or removes the modules in Load Modules for you">{features}</menu>
      {edit("loadmodules", "Load Modules", "a module; another module; another", "Python modules to load", "Comma separated. Must be importable from Python's search path")}
      {edit("searchpath", "Search Paths", "a module; another module; another", "Folders added to Python's search path", "Semi-colon separated folders holding your own modules")}
      <labelControl id="lb_restart" tag="restart" label="!Restart Required!" getVisible="xlo_GetVisible"/>
    </group>
    <group id="g_tools" label="Tools">
      {button("doctor", "Doctor", "ReviewProtectDocument", "Check the installation", "Checks bitness, runtime, packages, the add-in and settings, and says how to fix problems")}
      {button("reload", "Reload", "Refresh", "Reload Python modules", "Reloads your modules without restarting Excel")}
      {button("settings", "Settings File", "FileProperties", "Open xlOil.ini", "Most changes need an Excel restart")}
    </group>
    <group id="g_debug" label="Debugging">
      {button("console", "Console", "ContainerGallery", "Python console")}
      {button("qtconsole", "QtConsole", "ContainerGallery", "Jupyter console in Excel's Python")}
      {button("openlog", "Open Log", "ZoomCurrent75", "Open the log file")}
      {drop("loglevel", "Log Level", "Messages at or above this level are logged")}
      {drop("debugger", "Debugger", "Debugger used for Python code")}
      {edit("debugpyport", "DebugPy Port", "00000", "Port for the debugpy / VS Code debugger")}
    </group>
    <group id="g_other" label="Other">
      {edit("dateformats", "Date Formats", "%Y-%m-%d;%Y%b%d", "Formats tried when reading text as dates")}
      <checkBox id="cb_errorprop" tag="errorprop" label="Propagate Errors" getPressed="xlo_GetPressed" onAction="xlo_OnToggle" screentip="Return argument errors (e.g. #N/A) directly" supertip="Requires restart"/>
      {button("fixname", "Fix #NAME!", "ErrorChecking", "Recalculate #NAME! errors", size="normal")}
    </group>
  </tab></tabs></ribbon>
</customUI>'''


def _parts() -> dict:
    sys.path.insert(0, str(ROOT / "tools"))
    from vba_project import build_vba_project
    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    rels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkg_rels = "http://schemas.openxmlformats.org/package/2006/relationships"
    return {
        "[Content_Types].xml": xml +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.addin.macroEnabled.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '</Types>',
        "_rels/.rels": xml +
            f'<Relationships xmlns="{pkg_rels}">'
            f'<Relationship Id="rId1" Type="{rels}/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rIdXloRibbon" Type="http://schemas.microsoft.com/office/2006/relationships/'
            'ui/extensibility" Target="customUI/customUI.xml"/>'
            '</Relationships>',
        "xl/workbook.xml": xml +
            f'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{rels}">'
            '<workbookPr codeName="ThisWorkbook"/>'
            '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": xml +
            f'<Relationships xmlns="{pkg_rels}">'
            f'<Relationship Id="rId1" Type="{rels}/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
            'Target="vbaProject.bin"/>'
            '</Relationships>',
        "xl/worksheets/sheet1.xml": xml +
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetPr codeName="Sheet1"/><sheetData/></worksheet>',
        "xl/vbaProject.bin": build_vba_project({"xlOilRibbon": VBA}),
        "customUI/customUI.xml": ribbon_xml(),
    }


def build(out: Path = OUT):
    out.parent.mkdir(parents=True, exist_ok=True)
    fixed_time = (2024, 1, 1, 0, 0, 0)          # same bytes on every build
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in _parts().items():
            z.writestr(zipfile.ZipInfo(name, fixed_time), data, zipfile.ZIP_DEFLATED)
    print("Wrote", out)


if __name__ == "__main__":
    build(Path(sys.argv[1]) if len(sys.argv) > 1 else OUT)
