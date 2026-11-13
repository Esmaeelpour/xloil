
from ._core import *

from .logging import *

from .register import (
    Arg, 
    func,
    scan_module,
    register_functions,
    Caller,
    )

from .importer import (
    linked_workbook,
    source_addin,
    get_event_loop,
    import_functions
    )

from .com import (
    constants,
    PauseExcel,
    MacroContext,
    fix_name_errors
    )

from .importer import reload

from .formatting import (
    Formatter,
    DataFrameFormatter,
    ConditionalFormatter,
    DateFormatter,
    NumberFormatter,
    )

from .lru import lru_cache

from .tables import (
    Table,
    TableReturn,
    read_table,
    write_table,
    )

from .metadata import workbook_metadata, WorkbookMetadata

from .menus import context_menu, menu

from .rtd import RtdSimplePublisher

from .type_converters import (
    ExcelValue,
    Cache, 
    SingleValue,
    AllowRange,
    Array,
    FastArray,
    converter,
    returner
    )

# Small hacky workaround for the jupyter connection feature.
# If we're imported in an ipython kernel and the variable 
# _xloil_jpy_impl exists, it means xloil has connected to tht
# kernel. In which case, fix up the `func` and `app` functions
# to do the same as before
try:
    ipy = get_ipython()
    
    impl = ipy.user_ns.get("_xloil_jpy_impl", None)
    if impl is not None:
        global func, app
        func = impl.func
        app = impl.app
except NameError:
    ... # Not in an ipython kernel

# Names which pull in heavier dependencies (plotting libraries, ctypes Win32 
# bindings) are loaded on first use
_LAZY_ATTRS = {
    'plot': ('xloil.plotting', 'plot'),
    'register_plot_type': ('xloil.plotting', 'register_plot_type'),
    'set_plot_mode': ('xloil.plotting', 'set_plot_mode'),
    'SheetControl': ('xloil.gui.sheet_control', 'SheetControl'),
    'place_control': ('xloil.gui.sheet_control', 'place_control'),
}

def __getattr__(name):
    target = _LAZY_ATTRS.get(name)
    if target is None:
        raise AttributeError(f"module 'xloil' has no attribute '{name}'")
    import importlib
    value = getattr(importlib.import_module(target[0]), target[1])
    globals()[name] = value
    return value
