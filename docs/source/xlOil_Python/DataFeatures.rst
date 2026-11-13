========================================
Tables, Formatting, Caching and Metadata
========================================

.. contents::
    :local:

Excel Tables
------------

*Excel Tables* (ListObjects) are first-class arguments and return values.

Reading a table
~~~~~~~~~~~~~~~

Annotate an argument with ``xloil.Table``. Callers can pass the table name as
text or any reference into the table. The function receives a pandas
DataFrame, or a dict of column arrays without pandas or with
``xlo.Table(pandas=False)``:

.. code-block:: python

    @xlo.func
    def total(tbl: xlo.Table, column: str) -> float:
        return float(tbl[column].sum())

.. code-block:: text

    =total("SalesTable", "Amount")
    =total(SalesTable[#All], "Amount")

Unlike a plain range argument, the table is read by name, so the function sees
rows added later without the formula changing.

Returning a table
~~~~~~~~~~~~~~~~~

Excel does not let formulas spill into tables. Annotate the return value with
``xlo.TableReturn(...)`` and xlOil writes the result as a real table just after
the calculation, below the calling cell by default. The cell shows a label:

.. code-block:: python

    @xlo.func
    def trades(n: int) -> xlo.TableReturn(name="Trades", style="TableStyleMedium9"):
        return load_trades(n)           # pandas or polars DataFrame, dict or 2-d array

``name``
    Table name. The default is derived from the calling cell.
``style``
    Excel table style.
``offset``
    (rows, columns) from the caller. The default ``(1, 0)`` is the row below.
``index``
    Write a pandas index as columns.
``label``
    Text for the calling cell. ``{name}`` and ``{rows}`` are substituted.

On recalculation the same table is resized and refilled, and cells it no
longer covers are cleared. Formulas outside the table that refer to it by
name, such as ``=SUM(Trades[Amount])``, keep working.

From macros
~~~~~~~~~~~

.. code-block:: python

    xlo.write_table(df, "Report!B2", name="Results", style="TableStyleLight9", show_totals=True)
    df = xlo.read_table("Results")

Polars
------

``import xloil.polars`` registers converters for ``polars.DataFrame``,
``polars.Series`` and (return only) ``polars.LazyFrame``:

.. code-block:: python

    import xloil as xlo
    import xloil.polars
    import polars as pl

    @xlo.func
    def by_region(df: pl.DataFrame) -> pl.DataFrame:
        return df.group_by("region").agg(pl.col("sales").sum()).sort("region")

    @xlo.func
    def lazy(df: pl.DataFrame):
        return df.lazy().filter(pl.col("sales") > 0)    # collected automatically

``xlo.PolarsFrame(headings=True, dates=None, infer=True)``
    ``headings``: the first row holds column names (when reading) or column
    names are written (when returning). ``dates``: columns holding Excel serial
    dates to convert to ``Datetime``. ``infer``: let polars pick strict column
    types; otherwise mixed columns become ``Object``.
``xlo.PolarsSeries(name="", headings=False)``
    One column (or row) as a Series.

``Date`` and ``Datetime`` columns are written as Excel dates. Timezone-aware
datetimes are written as wall-clock time in their own timezone. Nulls become
empty cells. ``Time``, ``Duration`` and ``Categorical`` values are written as
text.

Cell formatting from functions
------------------------------

Pass a *formatter* to ``xloil.func``. After the calculation, xlOil formats the
cells the result occupies (spilled arrays and DataFrames included):

.. code-block:: python

    @xlo.func(formatter=xlo.NumberFormatter(2, percent=True))
    def growth(old: float, new: float):
        return new / old - 1

    @xlo.func(formatter=xlo.DataFrameFormatter(
                columns={'Price': xlo.NumberFormatter(2)},
                alternate_row_color="#F2F2F2"))
    def prices(n: int):
        return make_prices(n)

    red_negatives = xlo.ConditionalFormatter([
        (lambda v: isinstance(v, float) and v < 0, xlo.Formatter(font_color="red", bold=True)),
    ])

    @xlo.func(formatter=red_negatives)
    def pnl(x: xlo.Array(float)):
        return x

Formatters
~~~~~~~~~~

``Formatter(...)``
    ``number_format``, ``bold``, ``italic``, ``underline``, ``font_name``,
    ``font_size``, ``font_color``, ``interior_color``, ``horizontal_alignment``
    (left/center/right/...), ``vertical_alignment``, ``wrap_text``, ``borders``
    (``True`` or a colour), ``auto_fit`` and ``clear`` (clear old formats first).
    Colours can be ``"#RRGGBB"``, a name, ``(r, g, b)`` or an Excel colour integer.
``DataFrameFormatter(header=, index=, values=, columns={}, alternate_row_color=, index_included=)``
    Styles the heading row(s), the index column(s), the body and individual columns.
    Works with pandas and polars.
``NumberFormatter(decimals=2, thousands=True, percent=False)``, ``DateFormatter("yyyy-mm-dd")``
    Shortcuts for common number formats.
``ConditionalFormatter(rules, default=None)``
    Picks a formatter per cell from ``(predicate, formatter)`` rules. It
    evaluates each cell, so use Excel's own conditional formatting for large
    ranges.

Formatters can also be applied directly from macros:
``xlo.Formatter(bold=True).apply("Sheet1!A1:C1")``.

Subclass ``Formatter`` and override ``apply_value(self, com_range, value)`` to
write your own. Formatters are not supported for async or RTD functions.

Saving cached objects in the workbook
-------------------------------------

xlOil's object cache normally lives only in memory, so cache references in a
re-opened workbook point at nothing until their functions recalculate. With
``xloil.cache_persist``, picklable objects are saved inside the workbook and
put back under the *same reference strings* when it is opened, so downstream
formulas work straight away:

.. code-block:: python

    import xloil.cache_persist
    xloil.cache_persist.enable()        # once, e.g. in a startup script

    @xlo.func(persist=True)
    def fit(x: xlo.Array(float), y: xlo.Array(float)):
        return ExpensiveModel().fit(x, y)       # returned as a cache reference

Choose what is saved:

* ``@func(persist=True)``: results of this function.
* ``cache_persist.persist_type(MyModel)``: every cached object of these types.
* ``cache_persist.mark_persistent(obj)``: a specific object, e.g. one added with
  ``xlo.cache.add``.
* ``enable(persist_all=True)``: everything picklable.

``enable()`` options: ``dumps``/``loads`` (e.g. ``cloudpickle``),
``trusted_only`` (do not restore into Protected View workbooks; default
``True``) and ``max_bytes`` (size limit per workbook; default 200 MB).
``save(wb)`` and ``restore(wb)`` can also be called manually.

Objects are stored as base64-encoded pickles in a Custom XML Part, which Excel
keeps but does not display. Unpickling can run arbitrary code, so only enable
restoring for workbooks you trust.

LRU cache for worksheet functions
---------------------------------

``functools.lru_cache`` fails on the unhashable values Excel passes (numpy
arrays, lists, dicts). ``xloil.lru_cache`` builds the key from the values:

.. code-block:: python

    @xlo.func
    @xlo.lru_cache(maxsize=256, ttl=60)
    def risk(positions: xlo.Array(float), vol: float = 0.2):
        return slow_risk_model(positions, vol)

    risk.cache_info()    # CacheInfo(hits=..., misses=..., maxsize=256, currsize=...)
    risk.cache_clear()

* Put ``@xlo.func`` above ``@xlo.lru_cache`` so xlOil sees the signature.
* Keys are typed: ``1`` and ``1.0`` are different entries.
* numpy arrays are keyed by dtype, shape and bytes; pandas objects by
  ``hash_pandas_object``; other objects by hash or pickle. If no key can be
  made, the call runs uncached.
* ``ttl`` (seconds) expires entries.
* Results are shared, so don't mutate a returned object.

Workbook metadata
-----------------

.. code-block:: python

    meta = xlo.workbook_metadata()             # active workbook, or pass a name
    meta.properties['Title'] = "Q3 Risk Report"
    meta.properties['Author']
    meta.custom['ModelVersion'] = "2.4.1"       # File > Info > Properties > Custom
    meta.data['last_run'] = {'user': 'me', 'rows': 1024}    # JSON store
    meta.to_dict()

``properties``
    Built-in document properties (Title, Subject, Author, Keywords, Comments,
    Category, Company, Manager, ...).
``custom``
    Custom document properties. Values may be str, int, float, bool or datetime.
``data``
    A dict of JSON values stored in a Custom XML Part, invisible to users and
    saved with the file.

These use the COM API, so call them from macros, commands, event handlers or
macro-type worksheet functions (the ``xloil.func`` default). Save the
workbook to keep changes.
