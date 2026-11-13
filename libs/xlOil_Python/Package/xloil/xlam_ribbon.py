"""
    The .xlam front-end for xlOil's settings ribbon.

    Some organisations' Office policy ("Require application add-ins to be signed by
    Trusted Publisher") blocks the unsigned COM add-in that `xloil.xloil_ribbon`
    uses to draw the *xlOil Py* tab. An ordinary ``.xlam`` workbook add-in in the
    XLSTART folder isn't affected, so xlOil ships ``xlOil.xlam``: a ribbon whose
    VBA forwards every button, box and menu to `xloRibbonCall` below, which runs
    the same logic as the COM ribbon (`xloil.ribbon_core`).

    Load this module (instead of ``xloil.xloil_ribbon``) and put ``xlOil.xlam`` in
    XLSTART; ``xloil ribbon xlam`` does both.
"""
import asyncio
import inspect
from pathlib import Path

import xloil
from xloil import ribbon_core as core

XLAM_NAME = "xlOil.xlam"


def xlam_path() -> Path:
    """ The xlOil.xlam shipped inside this package """
    return Path(__file__).with_name("data") / XLAM_NAME


class _Control:
    """ Stands in for Office's IRibbonControl: the callbacks only read id and tag """
    def __init__(self, tag: str):
        self.id = tag
        self.tag = tag


def _feature(kind):
    def call(tag, *args):
        return {
            "pressed": core.get_feature_pressed,
            "enabled": core.get_feature_enabled,
            "supertip": core.get_feature_supertip,
            "toggle": core.set_feature,
        }[kind](_Control(tag), *args)
    return call


# (kind, tag) -> callback(ctrl, *args). Only these can be reached from the .xlam.
_CALLBACKS = {
    ("count", "env"): core.get_python_environment_count,
    ("itemlabel", "env"): lambda c, i: core.get_python_environment(c, int(i)),
    ("selected", "env"): core.get_python_environment_selected,
    ("select", "env"): lambda c, item_id, i: core.set_python_environment(c, item_id, int(i)),
    ("text", "pyexe"): core.get_python_home,
    ("change", "pyexe"): core.set_python_home,
    ("text", "pypath"): core.get_python_path,
    ("change", "pypath"): core.set_python_path,
    ("text", "loadmodules"): core.get_load_modules,
    ("change", "loadmodules"): core.set_load_modules,
    ("text", "searchpath"): core.get_user_search_path,
    ("change", "searchpath"): core.set_user_search_path,
    ("visible", "restart"): core.get_restart_label_visible,
    ("action", "doctor"): core.press_doctor,
    ("action", "reload"): core.press_reload,
    ("action", "settings"): core.press_edit_settings,
    ("action", "console"): core.press_open_console,
    ("action", "qtconsole"): core.press_open_qtconsole,
    ("action", "openlog"): core.press_open_log,
    ("count", "loglevel"): core.get_log_level_count,
    ("itemlabel", "loglevel"): lambda c, i: core.get_log_level(c, int(i)),
    ("selected", "loglevel"): core.get_log_level_selected,
    ("select", "loglevel"): lambda c, item_id, i: core.set_log_level(c, item_id, int(i)),
    ("count", "debugger"): core.get_debugger_count,
    ("itemlabel", "debugger"): lambda c, i: core.get_debugger(c, int(i)),
    ("selected", "debugger"): core.get_debugger_selected,
    ("select", "debugger"): lambda c, item_id, i: core.set_debugger(c, item_id, int(i)),
    ("text", "debugpyport"): lambda c: str(core.get_debugpy_port(c)),
    ("change", "debugpyport"): core.set_debugpy_port,
    ("text", "dateformats"): core.get_date_formats,
    ("change", "dateformats"): core.set_date_formats,
    ("pressed", "errorprop"): core.get_error_propagation,
    ("toggle", "errorprop"): core.set_error_propagation,
    ("action", "fixname"): core._fix_name_errors,
}


def _run(result):
    """ Callbacks may be coroutines: run them to completion on a private loop """
    if inspect.isawaitable(result):
        loop = asyncio.new_event_loop()
        try:
            return loop.run_until_complete(result)
        finally:
            loop.close()
    return result


def call(kind: str, tag: str, *args):
    """ Dispatches one ribbon callback. Returns the value for getters """
    if tag.startswith("feature:"):
        return _run(_feature(kind)(tag[len("feature:"):], *args))
    fn = _CALLBACKS.get((kind, tag))
    if fn is None:
        raise KeyError(f"No xlOil ribbon callback for {kind}/{tag}")
    return _run(fn(_Control(tag), *args))


@xloil.func(macro=True, local=False,
            help="Internal: called by xlOil.xlam's ribbon. Not for use in cells.")
def xloRibbonCall(kind: str, tag: str, arg1=None, arg2=None):
    args = [a for a in (arg1, arg2) if a is not None]
    try:
        result = call(kind, tag, *args)
    except Exception:
        xloil.log_except(f"xlOil ribbon: {kind}/{tag} failed")
        return None
    return "" if result is None else result
