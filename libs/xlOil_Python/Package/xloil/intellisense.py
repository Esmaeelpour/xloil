"""
    In-sheet function tooltips via Excel-DNA IntelliSense.

    Excel only shows argument help for user-defined functions in the *Insert
    Function* wizard. The open-source `Excel-DNA IntelliSense
    <https://github.com/Excel-DNA/IntelliSense>`_ add-in adds the same pop-up
    tooltips built-in functions get as you type a formula. It reads descriptions
    for any add-in from a file next to it called ``<AddinName>.intellisense.xml``.

    This module writes that file from the help and argument descriptions you
    already give xlOil (docstrings, ``help=``, ``args=``)::

        import xloil.intellisense
        xloil.intellisense.enable()      # e.g. in LoadModules or a startup script

    `enable` writes the file for every loaded xlOil add-in now and again whenever
    functions are (re)registered. Workbook-local functions are described in a
    Custom XML Part inside their workbook, which IntelliSense also reads.

    You need the IntelliSense add-in loaded in Excel. Install it with::

        xloil intellisense install

    which downloads ``ExcelDna.IntelliSense64.xll`` (or the 32-bit version) from
    the Excel-DNA GitHub releases into Excel's XLSTART folder. IntelliSense
    picks up description files when it starts or when an add-in is loaded, so
    after the first `enable` restart Excel once; afterwards new functions in a
    reloaded module appear when the add-in is next loaded.
"""

import os
from pathlib import Path
from typing import Iterable, List, Optional
from xml.sax.saxutils import quoteattr

from .logging import log, log_except

NAMESPACE = "http://schemas.excel-dna.net/intellisense/1.0"
INTELLISENSE_RELEASE_URL = \
    "https://github.com/Excel-DNA/IntelliSense/releases/latest/download/{name}"
_MAX_DESCRIPTION = 255


def _clean(text: str) -> str:
    text = " ".join((text or "").split())
    return text[:_MAX_DESCRIPTION]


def _first_paragraph(doc: str) -> str:
    """ The first paragraph of a docstring: what Excel can reasonably display """
    if not doc:
        return ""
    paragraphs = [p.strip() for p in doc.strip().split("\n\n")]
    return paragraphs[0] if paragraphs else ""


def function_xml(name: str, description: str, args: Iterable[tuple], help_topic: str = "") -> str:
    """ Builds one <Function> element. `args` is an iterable of (name, description) """
    attrs = f"Name={quoteattr(name)} Description={quoteattr(_clean(_first_paragraph(description)))}"
    if help_topic:
        attrs += f" HelpTopic={quoteattr(help_topic)}"
    arg_xml = "".join(
        f"<Argument Name={quoteattr(a)} Description={quoteattr(_clean(d))} />"
        for a, d in args)
    return f"<Function {attrs}>{arg_xml}</Function>"


def document_xml(functions: List[str]) -> str:
    """ Wraps <Function> elements in the IntelliSense document """
    return (f'<IntelliSense xmlns="{NAMESPACE}"><FunctionInfo>'
            + "".join(functions) + "</FunctionInfo></IntelliSense>")


def spec_to_xml(spec) -> Optional[str]:
    """ Converts an xlOil `_FuncSpec` to a <Function> element (None for commands) """
    try:
        if "command" in (getattr(spec, "features", "") or ""):
            return None
    except Exception:
        pass
    args = []
    for arg in spec.args:
        flags = getattr(arg, "flags", "") or ""
        if "keywords" in flags:
            args.append((arg.name, arg.help or "Two-column range of keyword/value pairs"))
        elif "vargs" in flags:
            args.append((arg.name, arg.help or "Additional arguments"))
        else:
            args.append((arg.name, arg.help or ""))
    return function_xml(spec.name, spec.help or "", args)


def _addin_xml_path(addin) -> Path:
    xll = Path(addin.pathname)
    return xll.with_suffix(".intellisense.xml")


def write_addin_file(addin, extra_specs=()) -> Optional[Path]:
    """
        Writes the ``.intellisense.xml`` file for an xlOil add-in, describing all
        its registered functions. Returns the path written.
    """
    elements = []
    for spec in list(addin.functions()) + list(extra_specs):
        xml = spec_to_xml(spec)
        if xml:
            elements.append(xml)
    path = _addin_xml_path(addin)
    try:
        text = document_xml(elements)
        if path.exists() and path.read_text(encoding="utf-8") == text:
            return path
        path.write_text(text, encoding="utf-8")
        log.info(f"IntelliSense: wrote {len(elements)} functions to {path}")
        return path
    except OSError as e:
        log.warn(f"IntelliSense: cannot write {path}: {e}")
        return None


def write_workbook_part(workbook, specs) -> int:
    """
        Stores IntelliSense descriptions for workbook-local functions in a
        Custom XML Part inside the workbook. Main thread only.
    """
    from .metadata import _workbook_com, _write_xml_part
    elements = [x for x in (spec_to_xml(s) for s in specs) if x]
    wb = _workbook_com(workbook)
    _write_xml_part(wb, NAMESPACE, document_xml(elements) if elements else None)
    return len(elements)


def refresh():
    """ Rewrites the IntelliSense files for all loaded xlOil add-ins """
    import xloil_core
    count = 0
    for path, addin in xloil_core.xloil_addins.items():
        try:
            if write_addin_file(addin):
                count += 1
        except Exception:
            log_except(f"IntelliSense: failed for add-in {path}")
    return count


_enabled = {'value': False}

def enable():
    """
        Writes IntelliSense files now and after every function registration.
        Safe to call more than once.
    """
    refresh()
    if _enabled['value']:
        return
    _enabled['value'] = True

    # Hook the registration function so files stay current when modules reload
    from . import register
    original = register._register_functions

    def hooked(funcs, module=None, addin=None, append=False):
        result = original(funcs, module, addin, append)
        try:
            if addin is not None:
                write_addin_file(addin)
            else:
                refresh()
        except Exception:
            log_except("IntelliSense: update after registration failed")
        return result

    hooked._xloil_original = original
    register._register_functions = hooked


def default_xll_name() -> str:
    import struct
    return "ExcelDna.IntelliSense64.xll" if struct.calcsize("P") == 8 else "ExcelDna.IntelliSense.xll"


def install_intellisense_addin(target_dir: str = None, xll_name: str = None,
                               url: str = None) -> Path:
    """
        Downloads the Excel-DNA IntelliSense add-in into `target_dir` (default:
        Excel's XLSTART folder) so Excel loads it on start. Used by
        ``xloil intellisense install``.
    """
    import urllib.request
    xll_name = xll_name or default_xll_name()
    if target_dir is None:
        target_dir = Path(os.getenv('APPDATA')) / "Microsoft" / "Excel" / "XLSTART"
    target = Path(target_dir) / xll_name
    target.parent.mkdir(parents=True, exist_ok=True)
    source = url or INTELLISENSE_RELEASE_URL.format(name=xll_name)
    print(f"Downloading {source}")
    with urllib.request.urlopen(source, timeout=60) as response:
        data = response.read()
    if not data.startswith(b"MZ"):
        raise RuntimeError(f"Download from {source} is not a Windows DLL/XLL")
    target.write_bytes(data)
    print(f"Installed {target}")
    return target
