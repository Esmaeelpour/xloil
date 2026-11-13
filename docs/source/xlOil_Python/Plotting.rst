============================
Plots and Interactive Charts
============================

.. contents::
    :local:

Upstream xlOil can return a *matplotlib* figure from a worksheet function, which
it inserts as a picture (see :doc:`ExternalPackages`).
The ``xloil.plotting`` module extends this to more libraries and adds
**interactive** charts which keep zoom, pan, hover and clickable legends.

Supported libraries
-------------------

=================  ==========================================  =================
Library            Returned object                             Shown as
=================  ==========================================  =================
matplotlib         ``Figure`` (via ``xloil.matplotlib``),       picture
                   ``Axes``
pandas             the ``Axes`` from ``df.plot(...)``           picture
seaborn            ``Axes``, ``FacetGrid``, ``JointGrid``,      picture
                   ``PairGrid``
plotly             ``Figure``, ``FigureWidget``                 interactive
bokeh              any model / layout (``figure``, ``row``...) interactive
altair             ``Chart`` and compound charts                interactive
anything else      register with ``register_plot_type``        either
=================  ==========================================  =================

Only installed libraries are registered, so importing ``xloil.plotting`` is
safe in any environment.

Returning a chart
-----------------

Load the module at start-up (add it to ``LoadModules`` in ``xlOil.ini``, or
import it in your own module), then return the chart object:

.. code-block:: python

    import xloil as xlo
    import xloil.plotting

    @xlo.func
    def pnl_chart(dates: xlo.Array(object), pnl: xlo.Array(float)):
        import plotly.graph_objects as go
        fig = go.Figure(go.Scatter(x=dates, y=pnl, mode="lines"))
        fig.update_layout(title="P&L", margin=dict(l=20, r=20, t=40, b=20))
        return fig

    @xlo.func
    def histogram(values: xlo.Array(float)):
        import pandas as pd
        return pd.Series(values).plot.hist(bins=30)     # matplotlib Axes -> picture

    @xlo.func
    def pairs(data: xlo.PDFrame):
        import seaborn as sns
        return sns.pairplot(data)                      # PairGrid -> picture

The calling cell shows a short label such as ``[plotly chart]``. Each
recalculation updates the same chart rather than adding another one.

Controlling size, position and mode
-----------------------------------

Return ``xloil.plot(...)`` instead of the bare figure to choose options:

.. code-block:: python

    @xlo.func
    def chart(x: xlo.Array(float), y: xlo.Array(float)):
        import altair as alt, pandas as pd
        c = alt.Chart(pd.DataFrame({'x': x, 'y': y})).mark_line().encode(x='x', y='y')
        return xlo.plot(c, size=(600, 350), offset=(1, 0), label="Trend")

``size``
    (width, height) in points. Interactive charts default to 480 x 320. Pictures
    default to the image size.
``offset``
    (rows, columns) from the calling cell to the chart's top-left corner. The
    default is ``(0, 1)`` (next column) for interactive charts, so the calling
    cell stays clickable, and ``(0, 0)`` for pictures, as in ``xloil.matplotlib``.
``interactive``
    ``True``/``False`` forces a mode; the default follows ``set_plot_mode``.
``label``
    The text returned to the cell.

``xloil.set_plot_mode('auto' | 'interactive' | 'static')`` chooses the default
for libraries which support both. ``auto`` (default) uses interactive charts
when Qt WebEngine is available and pictures otherwise.

Requirements for interactive charts
-----------------------------------

Interactive charts are HTML pages shown in a Qt WebEngine web view placed on
the sheet as a :doc:`sheet control <SheetControls>`. You need:

.. code-block:: bat

    pip install qtpy PySide6            # or: PyQt5 PyQtWebEngine

and ``xloil.gui.qtpy`` in ``LoadModules`` *before* ``xloil.plotting`` or any
other module which imports Qt:

.. code-block:: toml

    [xlOil_Python]
    LoadModules=["xloil.xloil_ribbon", "xloil.gui.qtpy", "xloil.plotting", "mymodule"]

Qt WebEngine has to be initialised before the Qt application starts;
``xloil.gui.qtpy`` does this automatically when WebEngine is installed (set
``XLOIL_QT_NO_WEBENGINE=1`` to skip it).

Charts are written to ``%TEMP%\xloil_plots`` with their JavaScript embedded
(plotly, bokeh). This keeps them working offline. Altair embeds Vega offline
when ``vl-convert-python`` is installed and otherwise loads it from a CDN.

Static fallbacks need each library's image exporter: *kaleido* for plotly,
*vl-convert-python* for altair, *selenium* and a web driver for bokeh.

When the workbook is opened without xlOil, interactive charts show as grey
placeholder rectangles. They are drawn again when their functions recalculate.

Adding another plotting library
-------------------------------

.. code-block:: python

    import xloil.plotting as xp
    from mylib import Chart

    xp.register_plot_type(
        Chart,
        to_html=lambda chart: chart.render_html(standalone=True),    # interactive
        to_png=lambda chart, filename: chart.save_png(filename),     # picture
        name="mylib")

After this, worksheet functions returning a ``Chart`` display it automatically.
Pass ``returner=False`` to only allow explicit ``xlo.plot(chart)`` calls.

Limitations
-----------

* An interactive chart is a live window over the grid. It is not printed, not
  copied with the sheet, and not drawn in Excel's print preview or page layout
  view. Use a picture for printed reports.
* Charts follow scrolling with a short delay (see :doc:`SheetControls`).
* Frozen panes: the chart tracks the scrolling pane, so place charts below and
  right of the freeze lines.
