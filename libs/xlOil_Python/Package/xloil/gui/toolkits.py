"""
    Extensible bridge between Python GUI toolkits and xlOil's windows.

    xlOil hosts Python windows in two places: custom task panes (docked side
    panels) and sheet controls (widgets floating on a worksheet, see
    `xloil.gui.sheet_control`). Both need the same few things from a toolkit:

      * which thread the toolkit must be driven from (an *executor*),
      * how to create a widget from what the user passed in,
      * the native window handle (HWND) of the widget, and
      * how to destroy it.

    `GuiToolkit` captures exactly that. xlOil registers toolkits for Qt (PyQt /
    PySide via *qtpy*), Tk, wxPython and raw HWNDs. To support another toolkit,
    subclass `GuiToolkit` and call `register_toolkit`; afterwards objects from that
    toolkit can be passed directly to `ExcelGUI.attach_pane` and to `SheetControl`::

        from xloil.gui.toolkits import GuiToolkit, register_toolkit

        class MyToolkit(GuiToolkit):
            name = "mytk"
            def handles(self, obj):
                return isinstance(obj, mytk.Window)
            def executor(self):
                return my_gui_thread_executor     # a concurrent.futures.Executor
            def create(self, obj):
                return obj if isinstance(obj, mytk.Window) else obj()
            def get_hwnd(self, widget):
                widget.show()
                return widget.native_handle(), True  # (hwnd, attach_as_parent)
            def destroy(self, widget):
                widget.close()

        register_toolkit(MyToolkit())

    All `GuiToolkit` methods except `handles` and `executor` are called on the
    toolkit's executor thread.
"""

import concurrent.futures as futures
import sys
import typing
from typing import List, Optional, Tuple

from xloil.gui import CustomTaskPane
from xloil.logging import log


class GuiToolkit:
    """
        Base class describing how xlOil drives a GUI toolkit. See module docs.
    """
    name: str = "toolkit"

    def handles(self, obj) -> bool:
        """ Returns True if `obj` (a widget, widget class or factory) belongs to this toolkit """
        raise NotImplementedError()

    def executor(self) -> futures.Executor:
        """ Returns the executor whose thread owns this toolkit's GUI """
        raise NotImplementedError()

    def create(self, obj):
        """ Returns a widget from `obj`, which may already be a widget, a class or a factory """
        return obj() if callable(obj) and isinstance(obj, type) else obj

    def get_hwnd(self, widget) -> Tuple[int, bool]:
        """
            Makes the widget's native window exist and returns (hwnd, as_parent).
            `as_parent` True means xlOil re-parents the window into the host, which
            is preferred; False means the window is kept positioned over the host
            instead (needed for toolkits which misbehave when re-parented, e.g. wx).
        """
        raise NotImplementedError()

    def is_running(self) -> bool:
        """ 
            True if this toolkit's GUI thread has been started. Used to pick a 
            toolkit for widget factory functions.
        """
        return False

    def prepare_for_sheet(self, widget):
        """ Optional hook: called before a widget is placed on a worksheet """
        pass

    def destroy(self, widget):
        """ Releases the widget """
        pass


_TOOLKITS: List[GuiToolkit] = []


def register_toolkit(toolkit: GuiToolkit, first: bool = True):
    """
        Registers a `GuiToolkit`. If `first` is True it is consulted before the
        built-in toolkits, so it can override how existing widget types are hosted.
    """
    _TOOLKITS[:] = [t for t in _TOOLKITS if t.name != toolkit.name]
    if first:
        _TOOLKITS.insert(0, toolkit)
    else:
        _TOOLKITS.append(toolkit)


def registered_toolkits() -> List[GuiToolkit]:
    return list(_TOOLKITS)


def find_toolkit(obj, allow_factory: bool = True) -> Optional[GuiToolkit]:
    """
        Returns the first registered toolkit which handles `obj`, or None.

        If `obj` is a plain function (a widget factory) no toolkit can tell 
        from the object itself what it will create. In that case, if exactly 
        one GUI toolkit is running (e.g. only `xloil.gui.qtpy` has been 
        imported) that toolkit is returned. Otherwise pass ``toolkit=`` 
        explicitly.
    """
    for toolkit in _TOOLKITS:
        try:
            if toolkit.handles(obj):
                return toolkit
        except Exception:
            continue

    if allow_factory and callable(obj) and not isinstance(obj, type):
        running = [t for t in _TOOLKITS if t.is_running()]
        if len(running) == 1:
            return running[0]
    return None


class ToolkitTaskPane(CustomTaskPane):
    """
        A `CustomTaskPane` which hosts a widget from any registered `GuiToolkit`.
        Created automatically by `ExcelGUI.attach_pane` when it is given an object
        a toolkit recognises.
    """
    def __init__(self, obj, toolkit: GuiToolkit = None):
        super().__init__()
        self._toolkit = toolkit or find_toolkit(obj)
        if self._toolkit is None:
            raise TypeError(f"No registered GUI toolkit handles {type(obj).__name__}")
        self._widget = self._toolkit.executor().submit(self._toolkit.create, obj).result()

    @property
    def widget(self):
        return self._widget

    @property
    def toolkit(self) -> GuiToolkit:
        return self._toolkit

    def _get_hwnd(self):
        return self._toolkit.executor().submit(self._toolkit.get_hwnd, self._widget)

    def on_destroy(self):
        super().on_destroy()
        self._toolkit.executor().submit(self._toolkit.destroy, self._widget)


#
# Built-in toolkits
#

class QtToolkit(GuiToolkit):
    """ Qt via *qtpy* (PyQt5/6, PySide2/6). Requires `xloil.gui.qtpy` to be imported first. """
    name = "qt"

    def handles(self, obj):
        if 'qtpy' not in sys.modules or 'xloil.gui.qtpy' not in sys.modules:
            return False
        from qtpy.QtWidgets import QWidget
        return isinstance(obj, QWidget) or (isinstance(obj, type) and issubclass(obj, QWidget))

    def executor(self):
        from xloil.gui.qtpy import Qt_thread
        return Qt_thread()

    def is_running(self):
        mod = sys.modules.get('xloil.gui.qtpy')
        return mod is not None and getattr(mod, '_Qt_thread', None) is not None

    def create(self, obj):
        from qtpy.QtWidgets import QWidget
        return obj if isinstance(obj, QWidget) else obj()  # class or factory

    def get_hwnd(self, widget):
        from qtpy.QtCore import Qt
        widget.setWindowFlags(Qt.FramelessWindowHint)
        widget.show()
        return int(widget.winId()), True

    def destroy(self, widget):
        widget.close()
        widget.deleteLater()


class TkToolkit(GuiToolkit):
    """ tkinter *Toplevel* windows. Requires `xloil.gui.tkinter` to be imported first. """
    name = "tk"

    def handles(self, obj):
        if 'xloil.gui.tkinter' not in sys.modules:
            return False
        import tkinter
        return isinstance(obj, tkinter.Toplevel) or \
            (isinstance(obj, type) and issubclass(obj, tkinter.Toplevel))

    def executor(self):
        from xloil.gui.tkinter import Tk_thread
        return Tk_thread()

    def is_running(self):
        mod = sys.modules.get('xloil.gui.tkinter')
        return mod is not None and getattr(mod, '_Tk_thread', None) is not None

    def create(self, obj):
        import tkinter
        return obj if isinstance(obj, tkinter.Toplevel) else obj()

    def get_hwnd(self, widget):
        from ctypes import windll
        widget.deiconify()
        widget.overrideredirect(True)
        widget.update_idletasks()
        return windll.user32.GetParent(widget.winfo_id()), True

    def destroy(self, widget):
        widget.destroy()


class WxToolkit(GuiToolkit):
    """ wxPython *Frame*. Requires `xloil.gui.wx` to be imported first. """
    name = "wx"

    def handles(self, obj):
        if 'wx' not in sys.modules or 'xloil.gui.wx' not in sys.modules:
            return False
        import wx
        return isinstance(obj, wx.Frame) or (isinstance(obj, type) and issubclass(obj, wx.Frame))

    def executor(self):
        from xloil.gui.wx import wx_thread
        return wx_thread()

    def is_running(self):
        return 'xloil.gui.wx' in sys.modules

    def create(self, obj):
        import wx
        return obj if isinstance(obj, wx.Frame) else obj()

    def get_hwnd(self, widget):
        widget.Show()
        # wx misbehaves when re-parented, so use the "shadowed" attachment
        return widget.GetHandle(), False

    def destroy(self, widget):
        widget.Destroy()


class HwndWindow(int):
    """
        Wraps a raw Win32 window handle so it can be hosted by xlOil, e.g. a
        window created by another library (pywebview, a native DLL, ...)::

            gui.attach_pane(HwndWindow(hwnd), name="External")

        Pass ``as_parent=False`` if the window's owner objects to re-parenting.
    """
    def __new__(cls, hwnd: int, as_parent: bool = True):
        obj = super().__new__(cls, hwnd)
        obj.as_parent = as_parent
        return obj


class _ImmediateExecutor(futures.Executor):
    def submit(self, fn, *args, **kwargs):
        future = futures.Future()
        try:
            future.set_result(fn(*args, **kwargs))
        except BaseException as e:
            future.set_exception(e)
        return future


class HwndToolkit(GuiToolkit):
    """ Hosts an existing native window given as `HwndWindow` """
    name = "hwnd"
    _executor = _ImmediateExecutor()

    def handles(self, obj):
        return isinstance(obj, HwndWindow)

    def executor(self):
        return self._executor

    def create(self, obj):
        return obj

    def get_hwnd(self, widget):
        from ctypes import windll
        windll.user32.ShowWindow(int(widget), 5) # SW_SHOW
        return int(widget), widget.as_parent

    def destroy(self, widget):
        from ctypes import windll
        windll.user32.ShowWindow(int(widget), 0) # SW_HIDE


for _tk in (QtToolkit(), TkToolkit(), WxToolkit(), HwndToolkit()):
    register_toolkit(_tk, first=False)
