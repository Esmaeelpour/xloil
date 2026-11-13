"""
    Menu functions and right-click context menu entries.

    **Context menus** - add an entry to Excel's right-click menu::

        import xloil as xlo

        @xlo.context_menu("Double values")
        def double_selection():
            for cell in xlo.app().Selection:
                cell.Value = cell.Value * 2

    **Menus** - add an item to a menu in the *Add-ins* ribbon tab, optionally
    in a sub-menu::

        @xlo.menu("Refresh prices", menu="My Tools")
        def refresh(): ...

        @xlo.menu("Export CSV", menu="My Tools", sub_menu="Export")
        def export(): ...

    The decorated function is registered as an xlOil *command* (a VBA-style macro)
    and the menu item runs it. The function takes no arguments; use
    ``xlo.app().Selection`` or `xloil.selection` to find the cells clicked on.
    Macro options (``disable_screen_updating`` etc.) can be passed to both
    decorators.

    Context menus available: "Cell" (default), "Row", "Column", "Ply" (sheet tabs),
    "List Range Popup" (inside Excel tables), "Shapes", or a list of these.

    Items are created as *temporary* command bar controls, so they disappear
    when Excel closes and are re-added when the module is loaded. Reloading the
    module replaces its items rather than duplicating them. For full control of
    the ribbon (including `<contextMenus>` in ribbon XML), use `xloil.ExcelGUI`.
"""

import inspect
import threading
from typing import Iterable, List, Optional, Union

import xloil_core
from .logging import log, log_except
from .register import func

_MSO_CONTROL_BUTTON = 1
_MSO_CONTROL_POPUP = 10
_TAG_PREFIX = "xloil:"
_MENU_BAR = "Worksheet Menu Bar"

_lock = threading.Lock()
_ITEMS = {}       # command name -> item spec dict
_pending = {'scheduled': False}
_hooked = {'bye': False}


def _schedule_build():
    """ (Re)creates all menu items on the main thread, batched """
    with _lock:
        if _pending['scheduled']:
            return
        _pending['scheduled'] = True

    def build():
        with _lock:
            _pending['scheduled'] = False
            items = list(_ITEMS.values())
        try:
            app = xloil_core.app().to_com()
        except Exception:
            log_except("menus: could not get Excel application")
            return
        for spec in items:
            try:
                _create_item(app, spec)
            except Exception:
                log_except(f"menus: failed to add menu item '{spec['caption']}'")

    if not _hooked['bye']:
        xloil_core.event.PyBye += remove_all
        _hooked['bye'] = True
    # Retry until COM is available (e.g. during start-up)
    xloil_core.excel_callback(build, wait=0, retry=500)


def _find_by_tag(controls, tag):
    for i in range(controls.Count, 0, -1):
        ctrl = controls.Item(i)
        try:
            if ctrl.Tag == tag:
                yield ctrl
        except Exception:
            continue


def _delete_item(app, spec):
    tag = _TAG_PREFIX + spec['command']
    for bar_name in spec['bars']:
        try:
            bar = app.CommandBars(bar_name)
        except Exception:
            continue
        for parent in _parents_to_search(bar.Controls, spec):
            for ctrl in list(_find_by_tag(parent, tag)):
                ctrl.Delete()
        _prune_empty_popups(bar.Controls, spec)


def _parents_to_search(controls, spec):
    """ Yields the control collections which may contain this spec's button """
    yield controls
    for name in [spec.get('menu'), spec.get('sub_menu')]:
        if not name:
            continue
        nxt = []
        for i in range(1, controls.Count + 1):
            ctrl = controls.Item(i)
            try:
                if ctrl.Type == _MSO_CONTROL_POPUP and ctrl.Caption == name:
                    nxt.append(ctrl.Controls)
            except Exception:
                continue
        for c in nxt:
            yield c
            controls = c


def _prune_empty_popups(controls, spec):
    """ Removes popups created by xlOil which have become empty """
    for i in range(controls.Count, 0, -1):
        ctrl = controls.Item(i)
        try:
            if ctrl.Type == _MSO_CONTROL_POPUP and str(ctrl.Tag).startswith(_TAG_PREFIX + "popup"):
                _prune_empty_popups(ctrl.Controls, spec)
                if ctrl.Controls.Count == 0:
                    ctrl.Delete()
        except Exception:
            continue


def _get_popup(controls, caption, before=None):
    for i in range(1, controls.Count + 1):
        ctrl = controls.Item(i)
        try:
            if ctrl.Type == _MSO_CONTROL_POPUP and ctrl.Caption == caption:
                return ctrl
        except Exception:
            continue
    kwargs = dict(Type=_MSO_CONTROL_POPUP, Temporary=True)
    if before is not None:
        kwargs['Before'] = before
    popup = controls.Add(**kwargs)
    popup.Caption = caption
    popup.Tag = _TAG_PREFIX + "popup:" + caption
    return popup


def _create_item(app, spec):
    _delete_item(app, spec) # Avoid duplicates on reload
    for bar_name in spec['bars']:
        bar = app.CommandBars(bar_name)
        controls = bar.Controls
        if spec.get('menu'):
            controls = _get_popup(controls, spec['menu']).Controls
        if spec.get('sub_menu'):
            controls = _get_popup(controls, spec['sub_menu']).Controls

        kwargs = dict(Type=_MSO_CONTROL_BUTTON, Temporary=True)
        if spec.get('position') is not None:
            kwargs['Before'] = spec['position']
        button = controls.Add(**kwargs)
        button.Caption = spec['caption']
        button.OnAction = spec['command']
        button.Tag = _TAG_PREFIX + spec['command']
        if spec.get('begin_group'):
            button.BeginGroup = True
        if spec.get('face_id') is not None:
            button.FaceId = spec['face_id']
        if spec.get('tooltip'):
            button.TooltipText = spec['tooltip']
        log.debug(f"menus: added '{spec['caption']}' to {bar_name}")


def _register(fn, caption, bars, menu=None, sub_menu=None, name=None,
              begin_group=False, face_id=None, tooltip=None, position=None,
              macro_options=None):
    command_name = name or f"xloMenu_{fn.__name__}"
    wrapped = func(fn, name=command_name, command=True, local=False,
                   help=tooltip or (fn.__doc__ or ""), **(macro_options or {}))
    spec = dict(
        command=command_name, caption=caption, bars=list(bars), menu=menu,
        sub_menu=sub_menu, begin_group=begin_group, face_id=face_id,
        tooltip=tooltip, position=position, module=getattr(fn, "__module__", None))
    with _lock:
        _ITEMS[command_name] = spec
    _schedule_build()
    return wrapped


def _extract_macro_options(kwargs):
    from .register import _MACRO_OPTION_NAMES
    unknown = set(kwargs) - set(_MACRO_OPTION_NAMES)
    if unknown:
        raise TypeError(f"Unexpected arguments: {', '.join(sorted(unknown))}")
    return kwargs


def context_menu(caption: str = None,
                 menu: Union[str, Iterable[str]] = "Cell",
                 sub_menu: str = None,
                 name: str = None,
                 begin_group: bool = False,
                 face_id: int = None,
                 position: int = None,
                 **macro_options):
    """
        Decorator: adds the function to Excel's right-click menu.

        Parameters
        ----------
        caption: str
            The menu text, default: the function name. Use '&' before a letter
            to give it an access key
        menu: str or list of str
            Which context menu(s): "Cell", "Row", "Column", "Ply",
            "List Range Popup", "Shapes", ...
        sub_menu: str
            Put the item in a sub-menu of this name
        name: str
            The registered command name (default "xloMenu_<function name>")
        begin_group: bool
            Draw a separator line above the item
        face_id: int
            An Office icon id (FaceId) to show next to the item
        position: int
            1-based position in the menu (default: at the end)
        **macro_options:
            disable_screen_updating, disable_alerts, disable_calculation,
            disable_events, restore_selection, restore_active_sheet
    """
    bars = [menu] if isinstance(menu, str) else list(menu)
    options = _extract_macro_options(macro_options)
    def decorate(fn):
        return _register(fn, caption or fn.__name__, bars, sub_menu=sub_menu, name=name,
                         begin_group=begin_group, face_id=face_id, position=position,
                         macro_options=options)
    return decorate


def menu(caption: str = None,
         menu: str = "xlOil",
         sub_menu: str = None,
         name: str = None,
         begin_group: bool = False,
         face_id: int = None,
         tooltip: str = None,
         **macro_options):
    """
        Decorator: adds the function as an item in a menu, shown in the
        *Add-ins* ribbon tab.

        Parameters
        ----------
        caption: str
            The menu item text, default: the function name
        menu: str
            Top level menu name (default "xlOil")
        sub_menu: str
            Optional sub-menu inside `menu`
        name, begin_group, face_id:
            As for `context_menu`
        tooltip: str
            Tool tip text (default: the function's docstring)
        **macro_options:
            As for `context_menu`
    """
    options = _extract_macro_options(macro_options)
    def decorate(fn):
        return _register(fn, caption or fn.__name__, [_MENU_BAR], menu=menu, sub_menu=sub_menu,
                         name=name, begin_group=begin_group, face_id=face_id,
                         tooltip=tooltip or (fn.__doc__ or "").strip().split("\n")[0],
                         macro_options=options)
    return decorate


def remove_menu_item(command_name: str):
    """ Removes a menu item added by `menu` or `context_menu`, by command name """
    with _lock:
        spec = _ITEMS.pop(command_name, None)
    if spec is None:
        return
    def work():
        _delete_item(xloil_core.app().to_com(), spec)
    xloil_core.excel_callback(work)


def remove_all():
    """ Removes every menu item added by xlOil """
    with _lock:
        specs = list(_ITEMS.values())
        _ITEMS.clear()
    try:
        app = xloil_core.app().to_com()
        for spec in specs:
            try:
                _delete_item(app, spec)
            except Exception:
                pass
    except Exception:
        pass


def menu_items() -> List[dict]:
    """ Returns descriptions of the menu items currently registered """
    with _lock:
        return [dict(spec) for spec in _ITEMS.values()]
