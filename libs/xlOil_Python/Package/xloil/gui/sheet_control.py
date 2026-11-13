"""
    Sheet controls: live Python widgets placed *on a worksheet*.

    A `SheetControl` puts a widget from any supported GUI toolkit (Qt, Tk, wx,
    a raw HWND, or a toolkit you register - see `xloil.gui.toolkits`) on top of a
    worksheet area. The widget moves with the cells when the sheet scrolls or
    zooms, hides when another sheet is shown and is removed when the sheet's
    placeholder shape is deleted. This gives the same experience as an ActiveX
    control, but with the full power of the Python GUI toolkit::

        import xloil as xlo
        from xloil.gui.qtpy import Qt_thread
        from xloil.gui.sheet_control import SheetControl

        @xlo.func(command=True)
        def add_slider():
            def make():
                from qtpy.QtWidgets import QSlider
                from qtpy.QtCore import Qt
                s = QSlider(Qt.Horizontal)
                s.valueChanged.connect(lambda v: write_value(v))
                return s
            SheetControl(make, at="B2", size=(200, 30), name="MySlider")

    How it works: xlOil adds a rectangle shape to the sheet as a placeholder
    (it stores the control's name and position in the workbook and is what users
    select, move, resize or delete). The widget's window is made a child of the
    Excel grid window and kept aligned with the placeholder. A light-weight timer
    on Excel's main thread checks the alignment (by default every 75ms), so the
    widget may trail the grid very slightly when scrolling quickly.

    The widget is a real window: it receives keyboard and mouse input directly.
    Since it's not an OLE object, it is not printed and does not appear when the
    workbook is opened without xlOil - the placeholder shape shows instead, with
    its text (``placeholder_text``) explaining what should be there.

    Sheet controls must be created on Excel's main thread: from a command/macro,
    an event handler or a ribbon callback. From a worksheet function, use
    `place_control`, which defers creation until after the calculation cycle.
"""

import ctypes
import ctypes.wintypes as wt
import itertools
import threading
import weakref
from typing import Callable, Dict, Optional, Tuple

import xloil_core
from xloil.logging import log, log_except
from .toolkits import find_toolkit, GuiToolkit

_user32 = ctypes.windll.user32 if hasattr(ctypes, "windll") else None

GWL_STYLE = -16
GWLP_HWNDPARENT = -8
WS_CHILD = 0x40000000
WS_POPUP = 0x80000000
WS_CAPTION = 0x00C00000
WS_THICKFRAME = 0x00040000
WS_CLIPCHILDREN = 0x02000000
WS_CLIPSIBLINGS = 0x04000000
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_NOZORDER = 0x0004
SWP_NOACTIVATE = 0x0010
SWP_SHOWWINDOW = 0x0040
SWP_HIDEWINDOW = 0x0080
SWP_ASYNCWINDOWPOS = 0x4000
SW_HIDE = 0
SW_SHOWNOACTIVATE = 4
HWND_TOP = 0

_MSO_SHAPE_RECTANGLE = 1
_XL_MOVE = 2   # Placement: move but don't size with cells

if _user32 is not None:
    _GetWindowLongPtr = getattr(_user32, "GetWindowLongPtrW", _user32.GetWindowLongW)
    _SetWindowLongPtr = getattr(_user32, "SetWindowLongPtrW", _user32.SetWindowLongW)
    _GetWindowLongPtr.restype = ctypes.c_ssize_t
    _GetWindowLongPtr.argtypes = (wt.HWND, ctypes.c_int)
    _SetWindowLongPtr.restype = ctypes.c_ssize_t
    _SetWindowLongPtr.argtypes = (wt.HWND, ctypes.c_int, ctypes.c_ssize_t)
    _user32.FindWindowExW.restype = wt.HWND
    _user32.FindWindowExW.argtypes = (wt.HWND, wt.HWND, wt.LPCWSTR, wt.LPCWSTR)
    _user32.SetParent.restype = wt.HWND
    _user32.SetParent.argtypes = (wt.HWND, wt.HWND)
    _user32.SetWindowPos.argtypes = (wt.HWND, wt.HWND, ctypes.c_int, ctypes.c_int,
                                     ctypes.c_int, ctypes.c_int, ctypes.c_uint)
    _user32.ScreenToClient.argtypes = (wt.HWND, ctypes.POINTER(wt.POINT))
    _user32.ShowWindowAsync.argtypes = (wt.HWND, ctypes.c_int)
    _user32.IsWindow.argtypes = (wt.HWND,)


def find_grid_window(xlmain_hwnd: int) -> int:
    """ Returns the EXCEL7 (worksheet grid) window below an XLMAIN window """
    desk = _user32.FindWindowExW(xlmain_hwnd, None, "XLDESK", None)
    if not desk:
        raise RuntimeError(f"No XLDESK window under {xlmain_hwnd:#x}")
    grid = _user32.FindWindowExW(desk, None, "EXCEL7", None)
    if not grid:
        raise RuntimeError(f"No EXCEL7 window under {xlmain_hwnd:#x}")
    return grid


def _resolve_sheet(sheet):
    """ Returns a COM Worksheet from None / name / '[Book]Sheet' / xloil or COM worksheet """
    app = xloil_core.app()
    if sheet is None:
        return app.to_com().ActiveSheet
    if isinstance(sheet, str):
        if sheet.startswith("["):
            book, name = sheet[1:].split("]", 1)
            return app.workbooks[book].worksheets[name].to_com()
        return xloil_core.active_workbook().worksheets[sheet].to_com()
    to_com = getattr(sheet, "to_com", None)
    return to_com() if to_com else sheet


_counter = itertools.count(1)
_CONTROLS: Dict[str, "SheetControl"] = {}
_lock = threading.RLock()


class SheetControl:
    """
        Places a GUI widget on a worksheet. Must be constructed on Excel's main
        thread (commands, macros, event handlers, ribbon callbacks).

        Parameters
        ----------
        widget:
            A widget, widget class or zero-argument factory recognised by a
            registered `GuiToolkit` (e.g. a Qt QWidget subclass). Factories are
            run on the toolkit's GUI thread.
        sheet:
            Worksheet name, '[Book]Sheet', `xloil.Worksheet` or COM worksheet.
            Default: the active sheet
        at: str or (left, top)
            Cell address for the top-left corner (e.g. "D4") or a position in points
        size: (width, height)
            Size in points. 'cell' sizes to the range given in `at`
        name: str
            Unique name, also used for the placeholder shape. If a control with
            this name exists on the sheet it is replaced, keeping its position
        placeholder_text: str
            Text shown in the placeholder shape when the widget is not present
            (e.g. the workbook is opened on a machine without xlOil)
        toolkit: GuiToolkit
            Force a toolkit rather than detecting one
        keep_position: bool
            If a placeholder called `name` already exists, keep its position and
            size rather than moving it to `at`/`size` (default True)
    """

    def __init__(self,
                 widget,
                 sheet=None,
                 at="B2",
                 size=(400, 300),
                 name: str = None,
                 placeholder_text: str = None,
                 toolkit: GuiToolkit = None,
                 keep_position: bool = True):

        self._toolkit = toolkit or find_toolkit(widget)
        if self._toolkit is None:
            raise TypeError(f"No registered GUI toolkit handles {type(widget).__name__}. "
                            "Import the toolkit's xloil.gui module (e.g. xloil.gui.qtpy) first.")

        self._sheet = _resolve_sheet(sheet)
        self._workbook_name = self._sheet.Parent.Name
        self._sheet_name = self._sheet.Name
        self.name = name or f"xloCtrl{next(_counter)}"
        self._key = f"[{self._workbook_name}]{self._sheet_name}!{self.name}"
        self._last_rect = None
        self._visible = False
        self._destroyed = False

        # Replace any live control of the same name
        with _lock:
            old = _CONTROLS.pop(self._key, None)
        if old is not None:
            old._release_widget()

        self._shape = self._find_or_create_shape(at, size, placeholder_text, keep_position)

        executor = self._toolkit.executor()
        self._widget = executor.submit(self._toolkit.create, widget).result()
        executor.submit(self._toolkit.prepare_for_sheet, self._widget).result()
        self._hwnd, self._as_parent = executor.submit(self._toolkit.get_hwnd, self._widget).result()
        self._hwnd = int(self._hwnd)

        self._window = self._find_window()
        self._xlmain = int(self._window.Hwnd)
        self._grid = find_grid_window(self._xlmain)
        self._attach()

        with _lock:
            _CONTROLS[self._key] = self
        _Tracker.instance().ensure_running()
        self.update_position()

    # ------------------------------------------------------------------ props
    @property
    def widget(self):
        """ The hosted widget. Interact with it only on its toolkit's thread """
        return self._widget

    @property
    def toolkit(self) -> GuiToolkit:
        return self._toolkit

    @property
    def shape(self):
        """ The COM Shape used as placeholder (main thread only) """
        return self._shape

    @property
    def hwnd(self) -> int:
        return self._hwnd

    @property
    def workbook(self) -> str:
        """ The workbook name (kept current across Save As / rename) """
        self._refresh_names()
        return self._workbook_name

    @property
    def sheet(self) -> str:
        """ The sheet name (kept current across renames) """
        self._refresh_names()
        return self._sheet_name

    def _refresh_names(self):
        """ Re-reads the workbook and sheet names: they change on Save As / rename """
        try:
            self._workbook_name = self._sheet.Parent.Name
            self._sheet_name = self._sheet.Name
        except Exception:
            pass

    def submit(self, fn, *args, **kwargs):
        """ Runs ``fn(widget, *args, **kwargs)`` on the toolkit thread, returning a Future """
        return self._toolkit.executor().submit(fn, self._widget, *args, **kwargs)

    # ---------------------------------------------------------------- set-up
    def _find_or_create_shape(self, at, size, placeholder_text, keep_position):
        shapes = self._sheet.Shapes
        shape = None
        try:
            shape = shapes.Item(self.name)
        except Exception:
            shape = None

        if isinstance(at, str):
            anchor = self._sheet.Range(at)
            left, top = float(anchor.Left), float(anchor.Top)
            if size == "cell":
                size = (float(anchor.Width), float(anchor.Height))
        else:
            left, top = at

        width, height = size

        if shape is None:
            shape = shapes.AddShape(_MSO_SHAPE_RECTANGLE, left, top, width, height)
            shape.Name = self.name
            shape.Placement = _XL_MOVE
            try:
                shape.Fill.ForeColor.RGB = 0xF2F2F2
                shape.Line.ForeColor.RGB = 0xBFBFBF
                shape.Line.Weight = 0.75
                frame = shape.TextFrame2.TextRange
                frame.Text = placeholder_text or f"Python control '{self.name}' (requires xlOil)"
                frame.Font.Size = 9
                frame.Font.Fill.ForeColor.RGB = 0x7F7F7F
            except Exception:
                pass
        elif not keep_position:
            shape.Left, shape.Top, shape.Width, shape.Height = left, top, width, height

        if placeholder_text is not None:
            try:
                shape.TextFrame2.TextRange.Text = placeholder_text
            except Exception:
                pass
        return shape

    def _find_window(self):
        wb = self._sheet.Parent
        app = wb.Application
        try:
            if app.ActiveWindow is not None and app.ActiveWindow.Parent.Name == wb.Name:
                return app.ActiveWindow
        except Exception:
            pass
        return wb.Windows(1)

    def _attach(self):
        style = _GetWindowLongPtr(self._hwnd, GWL_STYLE)
        self._previous_style = style
        if self._as_parent:
            style = (style | WS_CHILD | WS_CLIPSIBLINGS) & ~(WS_POPUP | WS_CAPTION | WS_THICKFRAME)
            _SetWindowLongPtr(self._hwnd, GWL_STYLE, style)
            self._previous_parent = _user32.SetParent(self._hwnd, self._grid)
            # Stop the grid painting over its children
            grid_style = _GetWindowLongPtr(self._grid, GWL_STYLE)
            if not grid_style & WS_CLIPCHILDREN:
                _SetWindowLongPtr(self._grid, GWL_STYLE, grid_style | WS_CLIPCHILDREN)
        else:
            style &= ~(WS_CAPTION | WS_THICKFRAME)
            _SetWindowLongPtr(self._hwnd, GWL_STYLE, style)
            # An owned top-level window stays above Excel and out of the taskbar
            _SetWindowLongPtr(self._hwnd, GWLP_HWNDPARENT, self._xlmain)
            self._previous_parent = None

    # -------------------------------------------------------------- tracking
    def _target_rect(self) -> Optional[Tuple[int, int, int, int]]:
        """ 
            Returns (x, y, w, h) in pixels relative to the grid window (or the
            screen if not attached as a child), or None if it should be hidden.
        """
        window = self._window
        self._refresh_names()
        if window.ActiveSheet.Name != self._sheet_name or \
                window.Parent.Name != self._workbook_name or \
                not window.Visible or window.WindowState == -4140: # xlMinimized
            return None

        shape = self._shape
        left, top = float(shape.Left), float(shape.Top)
        right, bottom = left + float(shape.Width), top + float(shape.Height)

        panes = window.Panes
        pane = panes.Item(panes.Count) # The scrollable pane when frozen/split
        x0 = int(pane.PointsToScreenPixelsX(left))
        y0 = int(pane.PointsToScreenPixelsY(top))
        width = max(1, int(pane.PointsToScreenPixelsX(right)) - x0)
        height = max(1, int(pane.PointsToScreenPixelsY(bottom)) - y0)

        if self._as_parent:
            pt = wt.POINT(x0, y0)
            _user32.ScreenToClient(self._grid, ctypes.byref(pt))
            x0, y0 = pt.x, pt.y
        return x0, y0, width, height

    def update_position(self):
        """ Re-aligns the widget with its placeholder. Main thread only. """
        if self._destroyed:
            return
        if not _user32.IsWindow(self._hwnd):
            self.delete(remove_shape=False)
            return
        try:
            rect = self._target_rect()
        except Exception as e:
            # Shape deleted by the user, or the workbook was closed
            if not self._shape_exists():
                log.info(f"SheetControl '{self.name}' placeholder removed, deleting control")
                self.delete(remove_shape=False)
            else:
                log.debug(f"SheetControl '{self.name}' position update skipped: {e}")
            return

        if rect is None:
            if self._visible:
                _user32.ShowWindowAsync(self._hwnd, SW_HIDE)
                self._visible = False
            return

        if rect != self._last_rect or not self._visible:
            flags = SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS | SWP_SHOWWINDOW
            if self._as_parent:
                flags |= SWP_NOZORDER
            _user32.SetWindowPos(self._hwnd, HWND_TOP, rect[0], rect[1], rect[2], rect[3], flags)
            self._last_rect = rect
            self._visible = True

    def _shape_exists(self) -> bool:
        try:
            self._sheet.Shapes.Item(self.name)
            return True
        except Exception:
            return False

    # -------------------------------------------------------------- clean-up
    def _release_widget(self):
        """ Detaches the window from Excel and destroys the widget """
        if self._destroyed:
            return
        self._destroyed = True
        hwnd = self._hwnd
        try:
            if _user32.IsWindow(hwnd):
                _user32.ShowWindowAsync(hwnd, SW_HIDE)
                if self._as_parent:
                    # Hand the window back before Excel destroys the grid window: some
                    # toolkits (Qt) abort if their windows are destroyed by others
                    _user32.SetParent(hwnd, None)
                    _SetWindowLongPtr(hwnd, GWL_STYLE, self._previous_style)
        except Exception:
            pass
        try:
            self._toolkit.executor().submit(self._toolkit.destroy, self._widget)
        except Exception:
            log_except(f"SheetControl '{self.name}': error destroying widget")

    def delete(self, remove_shape: bool = True):
        """ Removes the control. By default the placeholder shape is deleted too """
        with _lock:
            if _CONTROLS.get(self._key) is self:
                del _CONTROLS[self._key]
        self._release_widget()
        if remove_shape:
            try:
                self._sheet.Shapes.Item(self.name).Delete()
            except Exception:
                pass

    def __repr__(self):
        return f"SheetControl({self._key}, toolkit={self._toolkit.name})"




class _Tracker:
    """ Periodically re-aligns all sheet controls, on Excel's main thread """
    _instance = None
    interval_ms = 75

    @classmethod
    def instance(cls):
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        self._running = False
        self._hooked = False

    def ensure_running(self):
        if not self._hooked:
            xloil_core.event.WorkbookBeforeClose += _on_workbook_before_close
            xloil_core.event.PyBye += _on_bye
            self._hooked = True
        if not self._running:
            self._running = True
            self._schedule()

    def _schedule(self):
        xloil_core.excel_callback(self._tick, wait=self.interval_ms, retry=self.interval_ms)

    def _tick(self):
        with _lock:
            controls = list(_CONTROLS.values())
        for control in controls:
            try:
                control.update_position()
            except Exception:
                log_except(f"SheetControl tracker error for {control!r}")
        if controls:
            self._schedule()
        else:
            self._running = False


def set_tracking_interval(milliseconds: int):
    """ Sets how often (in ms) sheet controls are re-aligned with the grid. Default 75 """
    _Tracker.interval_ms = max(15, int(milliseconds))


def _on_workbook_before_close(wb_name, cancel):
    # Detach now: if the close is cancelled the tracker re-creates nothing, so
    # controls are only dropped on actual close. We can't know yet, so hide
    # windows; the WorkbookAfterClose handler releases them.
    # Match by window handle: names may have changed since the control was
    # created (Save As renames the workbook)
    hwnds = set()
    try:
        wb = xloil_core.app().to_com().Workbooks(wb_name)
        for i in range(1, wb.Windows.Count + 1):
            hwnds.add(int(wb.Windows(i).Hwnd))
    except Exception:
        pass
    with _lock:
        controls = [c for c in _CONTROLS.values()
                    if c._xlmain in hwnds or c.workbook == wb_name]
    for c in controls:
        try:
            # Reparent to desktop so Excel's window teardown doesn't destroy it
            if c._as_parent and _user32.IsWindow(c.hwnd):
                _user32.ShowWindowAsync(c.hwnd, SW_HIDE)
                _user32.SetParent(c.hwnd, None)
                c._as_parent_detached = True
        except Exception:
            pass
    def after():
        # If the workbook is still open, the close was cancelled: re-attach
        try:
            still_open = wb_name in xloil_core.app().workbooks
        except Exception:
            still_open = False
        for c in controls:
            if still_open and getattr(c, "_as_parent_detached", False):
                _user32.SetParent(c.hwnd, c._grid)
                c._as_parent_detached = False
                c._last_rect = None
            elif not still_open:
                c._shape = None
                c.delete(remove_shape=False)
    xloil_core.excel_callback(after, wait=250)


def _on_bye():
    with _lock:
        controls = list(_CONTROLS.values())
    for c in controls:
        try:
            c._release_widget()
        except Exception:
            pass
    _CONTROLS.clear()


def find_control(name: str, sheet=None) -> Optional[SheetControl]:
    """ Returns the live SheetControl called `name` on the given (or active) sheet """
    try:
        ws = _resolve_sheet(sheet)
        key = f"[{ws.Parent.Name}]{ws.Name}!{name}"
    except Exception:
        key = None
    with _lock:
        if key is not None and key in _CONTROLS:
            return _CONTROLS[key]
        return next((c for c in _CONTROLS.values() if c.name == name), None)


def all_controls():
    """ Returns a list of live sheet controls """
    with _lock:
        return list(_CONTROLS.values())


def place_control(widget, sheet=None, at="B2", size=(400, 300), name=None, **kwargs):
    """
        Like `SheetControl` but can be called from anywhere, including worksheet
        functions: creation is deferred to Excel's main thread after the current
        calculation. Returns a `concurrent.futures.Future` for the `SheetControl`.
    """
    import concurrent.futures
    result = concurrent.futures.Future()
    def create():
        try:
            result.set_result(SheetControl(widget, sheet=sheet, at=at, size=size, name=name, **kwargs))
        except BaseException as e:
            log_except("place_control failed")
            result.set_exception(e)
    xloil_core.excel_callback(create)
    return result
