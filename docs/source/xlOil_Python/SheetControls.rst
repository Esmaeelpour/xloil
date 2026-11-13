===============================
Sheet Controls and GUI Toolkits
===============================

.. contents::
    :local:

Sheet controls
--------------

A :py:class:`xloil.gui.sheet_control.SheetControl` places a widget from a Python
GUI toolkit *on the worksheet*, rather than in a side task pane. The widget
behaves like an ActiveX control: it sits over a cell area, moves when the sheet
scrolls or zooms, hides when you switch sheets, and takes keyboard and mouse
input directly.

.. code-block:: python

    import xloil as xlo
    import xloil.gui.qtpy                       # start Qt first
    from xloil.gui.sheet_control import SheetControl

    def make_slider():
        from qtpy.QtWidgets import QSlider
        from qtpy.QtCore import Qt
        slider = QSlider(Qt.Horizontal)
        slider.setRange(0, 100)

        def on_change(value):
            # Writing to the sheet must happen on Excel's main thread
            xlo.excel_callback(lambda: setattr(xlo.app().Range("C2"), "Value", value))
        slider.valueChanged.connect(on_change)
        return slider

    @xlo.func(command=True)
    def add_slider():
        SheetControl(make_slider, at="B2", size=(200, 24), name="RateSlider")

Run ``add_slider`` from the Macros dialog, a ribbon button or a
:doc:`menu <FunctionsAndMacros>`. Running it again replaces the widget and
keeps the position the user dragged it to.

Parameters
~~~~~~~~~~

``widget``
    A widget, widget class or zero-argument factory for any registered toolkit
    (Qt, Tk, wx, a raw window handle or your own). Factories run on the
    toolkit's GUI thread.
``sheet``
    Sheet name, ``"[Book]Sheet"``, ``xloil.Worksheet`` or COM worksheet. The
    default is the active sheet.
``at``
    A cell address for the top-left corner, or a ``(left, top)`` position in points.
``size``
    ``(width, height)`` in points, or ``"cell"`` to fit the range given in ``at``.
``name``
    A unique name, which is also the placeholder shape's name.
``placeholder_text``
    Text shown in the placeholder when xlOil is not loaded.
``keep_position``
    If a placeholder with this name exists, keep its position and size (the default).

Working with a control
~~~~~~~~~~~~~~~~~~~~~~

.. code-block:: python

    from xloil.gui.sheet_control import find_control

    ctrl = find_control("RateSlider")
    ctrl.submit(lambda w: w.setValue(50))      # runs on the Qt thread
    ctrl.delete()                              # removes widget and placeholder

* ``submit(fn, *args)`` runs ``fn(widget, *args)`` on the toolkit thread.
* ``delete(remove_shape=True)`` removes the control.
* ``all_controls()`` lists live controls.
* ``place_control(...)`` has the same arguments as ``SheetControl`` but can be
  called from anywhere, including worksheet functions. It defers creation to
  after the calculation and returns a ``Future``.
* ``set_tracking_interval(ms)`` changes how often controls realign (default 75 ms).

How it works
~~~~~~~~~~~~

xlOil adds a rectangle shape to the sheet as a placeholder. The placeholder
stores the control's name, position and size in the workbook, and it is what
users select, move, resize or delete. The widget's window becomes a child of
Excel's grid window (``EXCEL7``), so Windows clips it to the grid. A timer on
Excel's main thread keeps the window aligned with the placeholder using the
pane's ``PointsToScreenPixelsX/Y``.

Deleting the placeholder deletes the control. Closing the workbook releases the
widget before Excel destroys its windows; Qt in particular aborts if its
windows are destroyed by someone else.

Limitations
~~~~~~~~~~~

* Controls are live windows, not embedded OLE objects. They are not printed and
  are not saved in the file; the placeholder is. Re-create controls on
  workbook open, for example from a ``WorkbookOpen`` event handler or a
  workbook module.
* When scrolling quickly, the control can trail the grid by up to one tracking
  interval.
* With frozen or split panes, controls follow the bottom-right (scrolling) pane.
* While the control has keyboard focus, Excel shortcuts go to the control.

GUI toolkit bridge
------------------

Task panes (:doc:`CustomGUI`) and sheet controls host widgets through
:py:class:`xloil.gui.toolkits.GuiToolkit`. It describes how to drive a toolkit:

* ``handles(obj)``: does this object belong to the toolkit?
* ``executor()``: the ``concurrent.futures.Executor`` whose thread owns the GUI.
* ``create(obj)``: makes a widget from a widget, class or factory.
* ``get_hwnd(widget)``: returns ``(hwnd, as_parent)``. If ``as_parent`` is
  ``True``, xlOil re-parents the window into the host. If ``False``, xlOil keeps
  a top-level window positioned over the host instead, for toolkits that
  misbehave when re-parented.
* ``prepare_for_sheet(widget)``: an optional hook before sheet placement.
* ``destroy(widget)``: releases the widget.

Built-in toolkits: ``qt`` (``qtpy``: PyQt5/6, PySide2/6), ``tk``, ``wx`` and
``hwnd`` (wrap any existing window with ``xloil.gui.toolkits.HwndWindow(hwnd)``).

Registering a new toolkit makes its widgets usable with ``ExcelGUI.attach_pane``
and ``SheetControl``:

.. code-block:: python

    from xloil.gui.toolkits import GuiToolkit, register_toolkit
    import concurrent.futures

    class PyWebViewToolkit(GuiToolkit):
        name = "pywebview"
        _executor = concurrent.futures.ThreadPoolExecutor(1)

        def handles(self, obj):
            return type(obj).__name__ == "Window" and type(obj).__module__.startswith("webview")

        def executor(self):
            return self._executor

        def get_hwnd(self, window):
            return window.native.Handle.ToInt64(), True

        def destroy(self, window):
            window.destroy()

    register_toolkit(PyWebViewToolkit())

    gui.attach_pane(my_webview_window, name="Web")     # task pane
    SheetControl(my_webview_window, at="D2")           # on the sheet

Toolkits registered with ``first=True`` (the default) are checked before the
built-in ones, so they can change how an existing widget type is hosted.
