=======================
Fork Features: Overview
=======================

This fork of xlOil (`Esmaeelpour/xloil <https://github.com/Esmaeelpour/xloil>`_)
adds features which close the main gaps between xlOil and PyXLL. Everything
is implemented in the ``xloil`` Python package on top of xlOil's existing core,
so the compiled add-in is unchanged from upstream and existing code keeps
working. The features are opt-in: you import a module or pass a new decorator
argument.

.. contents::
    :local:

Feature map
-----------

.. list-table::
   :header-rows: 1
   :widths: 5 30 35 30

   * - #
     - Feature
     - How to use it
     - Docs
   * - 1
     - Interactive plots on the sheet (plotly, bokeh, altair)
     - ``import xloil.plotting`` then return the figure
     - :doc:`Plotting`
   * - 2
     - Wider plotting integrations (matplotlib, pandas, seaborn, plotly, bokeh, altair, custom)
     - ``xloil.plot(fig)``, ``xloil.register_plot_type``
     - :doc:`Plotting`
   * - 3
     - Widgets placed directly on the worksheet (ActiveX-style controls)
     - ``xloil.SheetControl(widget, at="B2")``
     - :doc:`SheetControls`
   * - 4
     - Extensible GUI toolkit bridge for task panes and sheet controls
     - ``xloil.gui.toolkits.register_toolkit``
     - :doc:`SheetControls`
   * - 5
     - Cached objects saved in the workbook and restored on open
     - ``xloil.cache_persist.enable()``, ``@func(persist=True)``
     - :doc:`DataFeatures`
   * - 6
     - Excel Tables as arguments and return values
     - ``xloil.Table``, ``xloil.TableReturn``, ``write_table``
     - :doc:`DataFeatures`
   * - 7
     - Polars DataFrame / Series conversion
     - ``import xloil.polars``
     - :doc:`DataFeatures`
   * - 8
     - RTD generators (plain ``def`` generators)
     - ``yield`` from an ``@func``
     - :doc:`FunctionsAndMacros`
   * - 9
     - Cell formatting from Python functions
     - ``@func(formatter=xloil.DataFrameFormatter())``
     - :doc:`DataFeatures`
   * - 10
     - Right-click context menu entries
     - ``@xloil.context_menu("Caption")``
     - :doc:`FunctionsAndMacros`
   * - 11
     - Menu functions (Add-ins tab menus)
     - ``@xloil.menu("Caption", menu="Tools")``
     - :doc:`FunctionsAndMacros`
   * - 12
     - Macro conveniences: screen updating, alerts, calculation, restore selection
     - ``@func(command=True, disable_screen_updating=True, ...)``
     - :doc:`FunctionsAndMacros`
   * - 13
     - LRU cache that accepts arrays, lists, dicts and DataFrames
     - ``@xloil.lru_cache(maxsize=128)``
     - :doc:`DataFeatures`
   * - 14
     - In-sheet IntelliSense tooltips (Excel-DNA IntelliSense)
     - ``xloil intellisense install`` + ``xloil.intellisense.enable()``
     - :doc:`DeveloperTools`
   * - 15
     - Reloading without restarting Excel
     - automatic on save (upstream), plus ``xloReload`` / ``xloil.reload()``
     - :doc:`DeveloperTools`
   * - 16
     - Command-line installer and configuration
     - ``xloil install --python ... --load-modules ...``, ``xloil config``, ``xloil status``
     - :doc:`DeveloperTools`
   * - 17
     - Startup scripts and virtual environments
     - ``StartupScript`` setting; ``xloil config python <venv>``
     - :doc:`DeveloperTools`
   * - 18
     - Workbook metadata (document properties and a JSON store)
     - ``xloil.workbook_metadata()``
     - :doc:`DataFeatures`
   * - 19
     - Testing stubs: import and unit test xlOil code without Excel
     - ``XLOIL_STUBS=1``, ``xloil.testing``, the ``xloil-stubs`` wheel
     - :doc:`DeveloperTools`
   * - 20
     - Stability and support
     - Policy, tests and CI in this fork
     - :doc:`Stability`

What already existed upstream
-----------------------------

Several items were already partly covered by xlOil, and this fork builds on them:

* **Reloading** (#15): xlOil watches loaded modules and reloads them, and
  re-registers their functions, when the file is saved. The fork adds an
  explicit reload function and commands.
* **Virtual environments** (#17): xlOil's loader already reads ``pyvenv.cfg``
  when ``PYTHONEXECUTABLE`` points into a venv. The fork adds the
  ``StartupScript`` setting and a CLI command to set the interpreter.
* **Command-line tool** (#16): ``xloil install/create/package/remove`` existed.
  The fork adds ``status``, ``config`` and ``intellisense`` and install options.
* **Async generators as RTD** (#8): ``async def`` generators already worked. The
  fork adds plain (synchronous) generators.
* **Stubs** (#19): xlOil shipped generated API stubs for documentation. The fork
  turns them into a working *stub mode* for tests.

Quick start
-----------

Full instructions, including conda, venvs and distributing to colleagues: :doc:`Installing`.

Install the wheel for your Python version from the GitHub release, then:

.. code-block:: bat

    pip install xloil-0.22.1.post5-cp312-cp312-win_amd64.whl
    xloil install --load-modules xloil.xloil_ribbon,xloil.plotting,mymodule

A module using several of the new features:

.. code-block:: python

    import xloil as xlo
    import xloil.polars          # registers polars converters
    import xloil.plotting        # registers chart return converters
    import polars as pl

    @xlo.func(formatter=xlo.DataFrameFormatter())
    @xlo.lru_cache(maxsize=32)
    def summary(data: pl.DataFrame) -> pl.DataFrame:
        return data.group_by("region").agg(pl.col("sales").sum())

    @xlo.func
    def sales_chart(data: pl.DataFrame):
        import plotly.express as px
        return px.bar(data.to_pandas(), x="region", y="sales")

    @xlo.func
    def clock():
        import time
        while True:
            yield time.strftime("%H:%M:%S")
            time.sleep(1)

    @xlo.context_menu("Clear formats", disable_screen_updating=True)
    def clear_formats():
        xlo.app().Selection.ClearFormats()
