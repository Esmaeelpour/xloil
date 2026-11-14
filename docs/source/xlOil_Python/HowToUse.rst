==================
How to Use xlOil
==================

This page starts where :doc:`Installing` ends. It takes you from an empty Excel
to your own Python functions in worksheet cells, then shows the everyday
tasks. Each section links to the full reference.

.. contents::
    :local:

1. Check that it works
----------------------

Start Excel. You should see an **xlOil Py** tab on the ribbon. Type these in
any cell:

=====================================  ============================================
Formula                                Result
=====================================  ============================================
``=xloVersion()``                      a small table with the xlOil version
``=xloAttr("python works","upper")``   ``PYTHON WORKS``, computed by Python
=====================================  ============================================

If you get ``#NAME?``, the add-in didn't load. Run ``xloil doctor`` at a command
prompt (in the Python environment you installed into) and fix what it reports.
More checks are in `7. When something goes wrong`_.

2. Write your first function
----------------------------

Make a folder for your code, for example ``C:\xl``, and create
``C:\xl\mytools.py``:

.. code-block:: python

    import xloil as xlo

    @xlo.func
    def Greeting(who: str) -> str:
        """Says hello"""
        return "Hello " + who

Tell xlOil where the file is and to load it. **Either** at a command prompt:

.. code-block:: bat

    xloil config add-path C:\xl
    xloil config add-module mytools

then restart Excel; **or** in Excel, on the *xlOil Py* tab, add ``C:\xl`` to
*Search Paths* (folders separated by ``;``) first, then add ``mytools`` to *Load
Modules* (modules separated by commas). Do it in that order: editing *Load
Modules* loads the new module straight away, so its folder must already be on
the search path.

Now type ``=Greeting("world")`` in a cell. The docstring appears as the
function's description in Excel's *Insert Function* dialog.

3. Change your code
-------------------

Edit ``mytools.py`` and save it: xlOil reloads the module and re-registers its
functions by itself. Press **F9** (or re-enter the formula) to see the new
result. If you change several files, or a module that others import, click
**Reload** on the *xlOil Py* tab.

4. Functions for one workbook only
----------------------------------

To keep functions with a single workbook, like a VBA module, put a Python file
with the workbook's name next to it: ``Budget.py`` beside ``Budget.xlsx``. xlOil
loads it when the workbook opens, and its functions are available in that
workbook only.

This needs *File > Options > Trust Center > Trust Center Settings > Macro
Settings > Trust access to the VBA project object model* to be ticked (it is off
by default). ``xloil install`` warns if it isn't. See :doc:`Concepts`.

5. Everyday recipes
-------------------

All of these go in a module like ``mytools.py``.

Arguments, defaults and help text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. code-block:: python

    @xlo.func(help="Price after discount",
              args={"price": "Original price", "pct": "Discount in percent"})
    def Discounted(price: float, pct: float = 10.0) -> float:
        return price * (1 - pct / 100)

``=Discounted(200)`` gives 180; ``=Discounted(200, 25)`` gives 150. Type hints
make xlOil convert the cell values, and report ``#VALUE!`` when they can't be
converted. More: :doc:`Functions`, :doc:`TypeConversion`.

Ranges and arrays
~~~~~~~~~~~~~~~~~

.. code-block:: python

    @xlo.func
    def Normalise(x: xlo.Array(float)):
        return x / x.sum()      # a numpy array: spills into the cells below

``=Normalise(A1:A5)`` takes a range as a numpy array. Returning an array (or a
list of lists) spills the result over the neighbouring cells.

Tables as pandas DataFrames
~~~~~~~~~~~~~~~~~~~~~~~~~~~

Needs the *pandas* feature (on by default; see *Features* on the ribbon).

.. code-block:: python

    import pandas as pd
    import xloil.pandas

    @xlo.func
    def Totals(sales: pd.DataFrame, by: str) -> pd.DataFrame:
        return sales.groupby(by, as_index=False).sum(numeric_only=True)

``=Totals(A1:C100, "region")`` reads the first row of the range as column
headings and writes the result back with headings. Options such as an index
column: :doc:`TypeConversion`. Excel Tables (``ListObject``) and polars:
:doc:`DataFeatures`.

Passing Python objects between cells
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

A function can return any Python object. A value Excel can't display (a model,
a dict of settings) shows in the cell as a reference to the object, and other
xlOil functions that take that cell get the original object:

.. code-block:: python

    class Model:
        def __init__(self, rate):
            self.rate = rate

    @xlo.func
    def MakeModel(rate: float):
        return Model(rate)

    @xlo.func
    def ModelValue(model, amount: float) -> float:
        return amount * (1 + model.rate)

``A1: =MakeModel(0.05)`` then ``=ModelValue(A1, 100)`` gives 105. The built-in
``=xloAttr(A1, "rate")`` reads any attribute or calls any method of such an
object. See :doc:`BuiltInUDFs`.

Values that update by themselves
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Write a generator: each ``yield`` updates the cell.

.. code-block:: python

    import time

    @xlo.func
    def Clock():
        while True:
            yield time.strftime("%H:%M:%S")
            time.sleep(1)

More, including ``async`` functions and RTD: :doc:`FunctionsAndMacros`, :doc:`Rtd`.

Charts
~~~~~~

Needs the *charts* feature (on by default). Return the figure through
``xlo.plot``; the chart appears next to the cell and is redrawn when the inputs
change:

.. code-block:: python

    @xlo.func
    def LineChart(y: xlo.Array(float)):
        import matplotlib.pyplot as plt
        fig, ax = plt.subplots()
        ax.plot(y)
        return xlo.plot(fig, size=(400, 250))

plotly, bokeh and altair figures give interactive charts. See :doc:`Plotting`.

Macros, buttons and right-click menus
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

A macro changes the workbook rather than returning a value:

.. code-block:: python

    @xlo.func(command=True)
    def FillDates():
        sel = xlo.app().Selection
        sel.Value = [[f"2026-01-{d:02d}"] for d in range(1, sel.Rows.Count + 1)]

    @xlo.context_menu("Clear formats")
    def ClearFormats():
        xlo.app().Selection.ClearFormats()

xlOil macros aren't listed in *View > Macros* (Alt+F8), but they run from
there: type ``FillDates`` as the macro name and click *Run*. You can also type
the name in a button's *Assign Macro* box.
*Clear formats* appears when you right-click a cell. More:
:doc:`FunctionsAndMacros`, :doc:`ExcelApplication`.

Reading and writing cells from Python
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Inside a macro, ``xlo.Range("Sheet1!A1:B2").value`` reads cells and assigning to
``.value`` writes them; ``xlo.app()`` is Excel's ``Application`` object, the same
one VBA uses. From a Python script or Jupyter outside Excel, ``xlo.Application()``
drives a running Excel. See :doc:`ExcelApplication`, :doc:`Jupyter`.

6. Turning features on and off
------------------------------

On the *xlOil Py* tab, the **Features** menu turns on pandas, polars, charts, Qt
widgets, IntelliSense tooltips and saving cached objects with the workbook.
Features whose libraries aren't installed are greyed out; the tooltip says what
to ``pip install``. Turning a feature off takes effect after restarting Excel.
To pick features again from a command prompt, run ``xloil setup``.

7. When something goes wrong
----------------------------

=========================================  ===================================================
What you see                               What to do
=========================================  ===================================================
``#NAME?`` for every xlOil function        The add-in didn't load. Run ``xloil doctor``.
``#NAME?`` for your function only          Check the *Search Paths* and *Load Modules* boxes,
                                           then **Open Log** for the import error (often a
                                           typo or a missing package).
A Python error message in the cell         Python raised an exception; the cell shows it
                                           with the traceback. **Open Log** has more.
Cells show ``#NAME?`` after reopening a    Click **Fix #NAME!** on the *xlOil Py* tab.
workbook
No *xlOil Py* tab                          Run ``xloil doctor``; see "Computers that require
                                           signed add-ins" in :doc:`Installing`.
=========================================  ===================================================

**Open Log** on the ribbon opens xlOil's log; ``=xloLog()`` returns its path.
**Console** opens a Python prompt running inside Excel, where you can import
your module and try it. To step through your code in a debugger, see
:doc:`Debugging`.

8. Command-line reference
-------------------------

Run these in the Python environment xlOil uses:

=================================  ============================================================
Command                            What it does
=================================  ============================================================
``xloil doctor``                   Checks the installation and says how to fix problems
``xloil status``                   Shows where xlOil is installed and how it is configured
``xloil setup``                    Guided set-up: choose features again, install their libraries
``xloil config add-module NAME``   Loads your module in Excel (``remove-module`` to stop)
``xloil config add-path FOLDER``   Adds a folder of your modules to the search path
``xloil config python PATH``       Points xlOil at another Python or venv
``xloil config list``              Prints every setting in ``xlOil.ini``
``xloil ribbon com|xlam|off``      Chooses how the *xlOil Py* tab is drawn
``xloil remove``                   Removes the add-in from Excel (keeps ``xlOil.ini``)
=================================  ============================================================

Changes made from the command line take effect when Excel restarts.

Where next
----------

* All the ways to declare functions: :doc:`Functions`
* Your own ribbon tabs and task panes: :doc:`CustomGUI`
* Widgets on the sheet: :doc:`SheetControls`
* Sharing your functions with colleagues: :doc:`DistributingAddins`
* Testing your code without Excel: :doc:`DeveloperTools`
