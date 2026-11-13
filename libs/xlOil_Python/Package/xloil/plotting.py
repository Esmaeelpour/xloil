"""
    Plotting integrations: return charts from worksheet functions.

    Import this module (or list it in ``LoadModules``) to register return
    converters for the plotting libraries you have installed::

        import xloil as xlo
        import xloil.plotting

        @xlo.func
        def price_chart(x: xlo.Array(float), y: xlo.Array(float)):
            import plotly.express as px
            return px.line(x=x, y=y)        # interactive chart next to the cell

    ===================  ====================================  ===============
    Library              Returned object                       Display
    ===================  ====================================  ===============
    matplotlib           ``Figure``, ``Axes``                  image
    pandas               ``df.plot()`` result (``Axes``)        image
    seaborn              ``Axes``, ``FacetGrid``, ``JointGrid`` image
    plotly               ``Figure`` / ``FigureWidget``          interactive
    bokeh                any layout / ``figure``                interactive
    altair               ``Chart`` (and compound charts)        interactive
    ===================  ====================================  ===============

    **Interactive** charts are rendered as HTML and shown in a web view placed on
    the sheet (an `xloil.gui.sheet_control.SheetControl`), so zoom, pan, hover
    and legends work exactly as in a browser. This needs Qt WebEngine::

        pip install qtpy PySide6       # or PyQt5 + PyQtWebEngine

    and ``xloil.gui.qtpy`` in ``LoadModules`` before ``xloil.plotting`` (or any
    other import of Qt). If WebEngine is unavailable, xlOil falls back to a
    static image when the library can export one (plotly needs *kaleido*, altair
    needs *vl-convert-python*, bokeh needs *selenium*).

    **Placement**: by default the chart is placed to the right of the calling
    cell and the cell shows a short label. Each recalculation updates the same
    chart in place. Use `plot` for control over size and position::

        @xlo.func
        def chart(...):
            return xlo.plot(fig, size=(600, 400), offset=(1, 0), interactive=True)

    **Other libraries**: register them with `register_plot_type`, giving a
    function to produce HTML (interactive) and/or one to write a PNG file.
"""

import hashlib
import os
import sys
import tempfile
import threading
from pathlib import Path
from typing import Callable, Dict, Optional, Tuple

import xloil_core
from .logging import log, log_except
from ._deferred import com_offset, com_resize, caller_location, run_after_calc, CellLocation
from .type_converters import return_converters, _make_typeconverter
from xloil_core import _CustomReturn

DEFAULT_SIZE = (480, 320)
_MODE = {'mode': 'auto'}   # auto | interactive | static
_TEMP_DIR = Path(tempfile.gettempdir()) / "xloil_plots"

class _PlotType:
    def __init__(self, to_html: Optional[Callable], to_png: Optional[Callable], name: str):
        self.to_html = to_html
        self.to_png = to_png
        self.name = name

_PLOT_TYPES: Dict[type, _PlotType] = {}


def set_plot_mode(mode: str):
    """
        Chooses how charts which support both modes are shown:

          * 'auto' (default): interactive if Qt WebEngine is available, else static
          * 'interactive': always interactive (error if WebEngine is missing)
          * 'static': always a picture
    """
    if mode not in ('auto', 'interactive', 'static'):
        raise ValueError("mode must be 'auto', 'interactive' or 'static'")
    _MODE['mode'] = mode


def webengine_available() -> bool:
    """ True if a Qt WebEngine web view can be created """
    if 'xloil.gui.qtpy' not in sys.modules:
        return False
    try:
        from qtpy import QtWebEngineWidgets  # noqa
        return True
    except Exception:
        return False


def _write_png_writer(png_func):
    """ Adapts a fn(obj, filename) into the writer insert_cell_image expects """
    def make(obj):
        return lambda filename: png_func(obj, filename)
    return make


def _location_name(location: CellLocation) -> str:
    digest = hashlib.md5(location.key().encode("utf-8")).hexdigest()[:10]
    return f"xloPlot_{digest}"


def _html_path(location: CellLocation) -> Path:
    _TEMP_DIR.mkdir(parents=True, exist_ok=True)
    return _TEMP_DIR / f"{_location_name(location)}.html"


class _WebViewFactory:
    """ Creates a QWebEngineView on the Qt thread showing a local html file """
    def __init__(self, path: Path):
        self.path = path

    def __call__(self):
        from qtpy.QtWebEngineWidgets import QWebEngineView
        from qtpy.QtCore import QUrl
        view = QWebEngineView()
        view.load(QUrl.fromLocalFile(str(self.path)))
        return view


def _reload_view(view, path):
    from qtpy.QtCore import QUrl
    view.load(QUrl.fromLocalFile(str(path)))


def _show_interactive(html: str, location: CellLocation, size, offset, name: str = None):
    from .gui.sheet_control import SheetControl, find_control
    from .gui.qtpy import Qt_thread  # noqa: ensures Qt is running
    from .gui.toolkits import find_toolkit

    path = _html_path(location)
    path.write_text(html, encoding="utf-8")
    name = name or _location_name(location)

    def place():
        sheet = f"[{location.workbook}]{location.sheet}"
        existing = find_control(name, sheet)
        if existing is not None:
            existing.submit(_reload_view, path)
            return
        anchor = com_offset(location.com_range(), offset[0], offset[1])
        from qtpy.QtWebEngineWidgets import QWebEngineView
        toolkit = find_toolkit(QWebEngineView)
        SheetControl(_WebViewFactory(path), sheet=sheet, at=anchor.Address, size=size,
                     name=name, toolkit=toolkit,
                     placeholder_text=f"Interactive chart (requires xlOil)")

    run_after_calc(place)


def _show_static(obj, png_func, location: CellLocation, size, offset):
    # insert_cell_image places relative to the caller and replaces any previous
    # image from the same caller.  Offset is converted to points from the caller.
    pos = (0, 0)
    origin = None
    if offset != (0, 0):
        def compute_pos():
            rng = location.com_range()
            target = com_offset(rng, offset[0], offset[1])
            return (float(target.Left - rng.Left), float(target.Top - rng.Top))
        try:
            pos = compute_pos()
        except Exception:
            pos = (0, 0)
    return xloil_core.insert_cell_image(lambda filename: png_func(obj, filename),
                                        size or "img", pos, origin)


def plot(obj, size: Tuple[float, float] = None, offset: Tuple[int, int] = None,
         interactive: Optional[bool] = None, name: str = None, label: str = None):
    """
        Displays a chart from a worksheet function. Call it in the return statement::

            return xlo.plot(fig, size=(500, 300))

        Parameters
        ----------
        obj:
            A chart from any registered plotting library
        size: (width, height)
            Size in points (default 480 x 320, or the image's size for static plots)
        offset: (rows, cols)
            Position of the chart's top-left relative to the calling cell.
            Default (0, 1): the next column on the same row
        interactive: bool
            Force interactive (True) or static (False). Default: `set_plot_mode`
        name: str
            Name for the chart's sheet control / placeholder shape
        label: str
            Text returned to the calling cell. Default "[<library> chart]"

        Returns the label, so the calling cell shows it.
    """
    handler = _find_plot_type(obj)
    if handler is None:
        raise TypeError(f"No plot handler for {type(obj).__name__}: use register_plot_type")

    if xloil_core.in_wizard():
        return f"[{handler.name} chart]"

    location = caller_location()
    if location is None:
        raise RuntimeError("plot() must be called from a worksheet function")

    if interactive is None:
        mode = _MODE['mode']
        if mode == 'static':
            interactive = False
        elif mode == 'interactive':
            interactive = True
        else:
            interactive = handler.to_html is not None and webengine_available()

    if offset is None:
        offset = (0, 1) if interactive else (0, 0)

    if interactive:
        if handler.to_html is None:
            raise TypeError(f"{handler.name} charts cannot be shown interactively")
        if not webengine_available():
            raise RuntimeError("Interactive charts need Qt WebEngine: install qtpy with PySide6 "
                               "(or PyQt5 + PyQtWebEngine) and load xloil.gui.qtpy at start-up")
        _show_interactive(handler.to_html(obj), location, size or DEFAULT_SIZE, offset, name)
    else:
        if handler.to_png is None:
            raise TypeError(f"{handler.name} charts cannot be exported as images. "
                            "Install the library's image export dependency")
        _show_static(obj, handler.to_png, location, size, offset)

    return label or f"[{handler.name} chart]"


def _find_plot_type(obj) -> Optional[_PlotType]:
    for typ, handler in _PLOT_TYPES.items():
        if isinstance(obj, typ):
            return handler
    return None


def register_plot_type(target: type,
                       to_html: Callable = None,
                       to_png: Callable = None,
                       name: str = None,
                       returner: bool = True):
    """
        Adds support for returning another plotting library's objects from
        worksheet functions.

        Parameters
        ----------
        target: type
            The chart class (instances and subclasses are matched)
        to_html: callable
            ``to_html(obj) -> str``: a complete standalone HTML document. Enables
            interactive display
        to_png: callable
            ``to_png(obj, filename)``: writes a PNG image. Enables static display
        name: str
            Library name used in labels and messages
        returner: bool
            If True (default), functions returning `target` objects display them
            automatically. If False, only explicit `plot(obj)` calls do
    """
    if to_html is None and to_png is None:
        raise ValueError("Specify at least one of to_html or to_png")
    _PLOT_TYPES[target] = _PlotType(to_html, to_png, name or target.__module__.split(".")[0])
    if not returner:
        return None

    def write(obj):
        return plot(obj)
    write.__name__ = f"ReturnPlot_{getattr(target, '__name__', 'plot')}"

    converter = _make_typeconverter(target, writer=_CustomReturn(write, write.__name__))
    return_converters.add(converter, target)
    return converter


#
# Built-in library support. Each block is skipped if the library is missing.
#

def _register_matplotlib():
    try:
        import matplotlib
    except ImportError:
        return
    # Use the non-interactive backend: we never want windows popping up in Excel
    if 'matplotlib.pyplot' not in sys.modules:
        matplotlib.use('Agg')
    from matplotlib.figure import Figure
    from matplotlib.axes import Axes

    def fig_png(fig, filename):
        fig.savefig(filename, format="png", bbox_inches="tight")

    def axes_png(ax, filename):
        fig_png(ax.get_figure(), filename)

    # Returning a Figure is already handled by xloil.matplotlib: don't change that
    # behaviour, but let `plot()` accept figures with its placement options
    register_plot_type(Figure, to_png=fig_png, name="matplotlib", returner=False)
    # Axes covers pandas `df.plot()` and seaborn axes-level functions
    register_plot_type(Axes, to_png=axes_png, name="matplotlib")

    # numpy arrays of Axes (from plt.subplots / df.plot(subplots=True)) are not
    # registered: they would capture every object array. Use plot(axes[0, 0]).


def _register_seaborn():
    try:
        import seaborn.axisgrid as sg
    except ImportError:
        return

    def grid_png(grid, filename):
        fig = getattr(grid, "figure", None) or getattr(grid, "fig")
        fig.savefig(filename, format="png", bbox_inches="tight")

    register_plot_type(sg.Grid, to_png=grid_png, name="seaborn")


def _register_plotly():
    try:
        from plotly.basedatatypes import BaseFigure
    except ImportError:
        return

    def to_html(fig):
        # Embed plotly.js so charts work offline (the file is local, not size limited)
        return fig.to_html(include_plotlyjs=True, full_html=True,
                           config={'responsive': True, 'displaylogo': False},
                           default_width="100%", default_height="100%")

    def to_png(fig, filename):
        fig.write_image(filename, format="png")  # requires kaleido

    register_plot_type(BaseFigure, to_html=to_html, to_png=to_png, name="plotly")


def _register_bokeh():
    try:
        from bokeh.model import Model
        from bokeh.embed import file_html
        from bokeh.resources import INLINE
    except ImportError:
        return

    def to_html(obj):
        try:
            obj.sizing_mode = "stretch_both"
        except Exception:
            pass
        return file_html(obj, INLINE, "xlOil chart")

    def to_png(obj, filename):
        from bokeh.io import export_png  # requires selenium + a webdriver
        export_png(obj, filename=filename)

    register_plot_type(Model, to_html=to_html, to_png=to_png, name="bokeh")


def _register_altair():
    try:
        import altair as alt
        base = alt.TopLevelMixin
    except (ImportError, AttributeError):
        return

    def to_html(chart):
        try:
            return chart.to_html(inline=True)  # offline, needs vl-convert-python
        except Exception:
            return chart.to_html()             # loads vega from a CDN

    def to_png(chart, filename):
        chart.save(filename, format="png")     # requires vl-convert-python

    register_plot_type(base, to_html=to_html, to_png=to_png, name="altair")


for _register in (_register_matplotlib, _register_seaborn, _register_plotly,
                  _register_bokeh, _register_altair):
    try:
        _register()
    except Exception:
        log_except(f"xloil.plotting: {_register.__name__} failed")
