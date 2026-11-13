"""
    Excel Tables (ListObjects) as a first-class feature.

    **Reading** - annotate an argument with `xloil.Table` and pass either a table
    name (as a string) or any reference into the table::

        @xlo.func
        def total_sales(tbl: xlo.Table) -> float:
            # tbl is a pandas DataFrame (or a dict of columns if pandas is missing)
            return tbl["Sales"].sum()

        # In Excel:  =total_sales("SalesTable")   or   =total_sales(SalesTable[#All])

    **Returning** - return a DataFrame with the `xloil.TableReturn` annotation to
    write it as a formatted Excel Table next to the calling cell. Excel does not
    allow formulas to spill into tables, so xlOil writes the table just after the
    calculation cycle (like PyXLL's table return). Each recalculation updates the
    same table, resizing it as needed::

        @xlo.func
        def load_trades(n: int) -> xlo.TableReturn(name="Trades", style="TableStyleMedium2"):
            return get_trades(n)

    **From macros** - `write_table` and `read_table` work with any range::

        xlo.write_table(df, "Sheet1!B2", name="Results")
        df = xlo.read_table("Results")
"""

from typing import Optional, Union, Any
import numpy as np

import xloil_core
from .type_converters import converter, returner
from ._core import CannotConvert, Range
from ._deferred import com_offset, com_resize, caller_location, run_after_calc, CellLocation
from .logging import log

_XL_SRC_RANGE = 1   # xlSrcRange
_XL_YES = 1         # xlYes (has headers)


def _find_list_object(name_or_range):
    """ Returns the COM ListObject for a table name, Range or COM range """
    app = xloil_core.app()
    if isinstance(name_or_range, str):
        name = name_or_range.split("[")[0].strip()
        # Table names are unique across the workbook; search all open books,
        # active workbook first
        books = [xloil_core.active_workbook()] + [b for b in app.workbooks]
        for wb in books:
            if wb is None:
                continue
            for ws in wb.worksheets:
                try:
                    return ws.to_com().ListObjects(name)
                except Exception:
                    continue
        # Maybe it's an address
        try:
            rng = app.to_com().Range(name_or_range)
        except Exception:
            raise CannotConvert(f"No table named '{name_or_range}'")
    else:
        to_com = getattr(name_or_range, "to_com", None)
        rng = to_com() if to_com else name_or_range

    lo = rng.ListObject
    if lo is None:
        raise CannotConvert(f"Range {rng.Address} is not inside a table")
    return lo


def _list_object_to_frame(lo, use_pandas=True):
    header = lo.HeaderRowRange
    headings = [str(h) for h in np.array(header.Value2, dtype=object).ravel()] if header else None
    body = lo.DataBodyRange
    ncols = lo.ListColumns.Count
    if headings is None:
        headings = [f"Column{i+1}" for i in range(ncols)]
    if body is None:
        data = np.empty((0, ncols), dtype=object)
    else:
        data = np.array(body.Value2, dtype=object).reshape(body.Rows.Count, body.Columns.Count)

    if use_pandas:
        try:
            import pandas as pd
            return pd.DataFrame(data, columns=headings).infer_objects()
        except ImportError:
            pass
    return {h: data[:, i] for i, h in enumerate(headings)}


def read_table(table, as_pandas: bool = True):
    """
        Reads an Excel table into a pandas DataFrame (or a dict of column arrays
        if `as_pandas` is False or pandas is not installed). `table` is a table
        name, an `xloil.Range` or a COM Range inside the table.
    """
    return _list_object_to_frame(_find_list_object(table), as_pandas)


@converter(range=True, check_cache=True)
class Table:
    """
        Argument converter: receives an Excel Table as a pandas DataFrame.
        The argument can be the table name or a reference to any part of it.
        Use ``xlo.Table(pandas=False)`` to get a dict of column arrays instead.
    """
    def __init__(self, pandas: bool = True):
        self._pandas = pandas

    def read(self, x):
        try:
            import pandas as pd
            if isinstance(x, pd.DataFrame):
                return x # Already converted, e.g. from cache
        except ImportError:
            pass
        if isinstance(x, (str, Range)):
            return read_table(x, self._pandas)
        raise CannotConvert(f"Expected a table name or range, got {type(x).__name__}")


def _frame_to_values(value, index: bool):
    """ Converts a frame-like value to (headings, 2d-object-array) """
    try:
        import pandas as pd
        if isinstance(value, pd.Series):
            value = value.to_frame()
        if isinstance(value, pd.DataFrame):
            frame = value.reset_index() if index else value
            headings = [str(c) if not isinstance(c, tuple) else " ".join(map(str, c))
                        for c in frame.columns]
            data = frame.to_numpy(dtype=object)
            # Excel can't take pandas NaT/Timestamp directly via COM: convert
            for j in range(data.shape[1]):
                col = frame.iloc[:, j]
                if pd.api.types.is_datetime64_any_dtype(col):
                    data[:, j] = [None if pd.isna(v) else v.to_pydatetime().replace(tzinfo=None)
                                  for v in col]
            data = np.where(pd.isna(data), None, data)
            return headings, data
    except ImportError:
        pass
    try:
        import polars as pl
        if isinstance(value, pl.DataFrame):
            return list(value.columns), np.array(value.rows(), dtype=object).reshape(value.height, value.width)
    except ImportError:
        pass
    if isinstance(value, dict):
        headings = [str(k) for k in value]
        cols = [np.asarray(v, dtype=object).ravel() for v in value.values()]
        n = max((len(c) for c in cols), default=0)
        data = np.full((n, len(cols)), None, dtype=object)
        for j, c in enumerate(cols):
            data[:len(c), j] = c
        return headings, data
    arr = np.asarray(value, dtype=object)
    if arr.ndim == 1:
        arr = arr.reshape(-1, 1)
    headings = [str(h) for h in arr[0]]
    return headings, arr[1:]


def _to_com_values(headings, data):
    """ Builds the nested tuples COM expects for Range.Value """
    import datetime as dt
    def clean(v):
        if isinstance(v, np.generic):
            v = v.item()
        if isinstance(v, float) and np.isnan(v):
            return None
        if isinstance(v, dt.datetime) and v.tzinfo is not None:
            v = v.replace(tzinfo=None)
        if v is not None and not isinstance(v, (str, int, float, bool, dt.date, dt.datetime)):
            return str(v)
        return v
    rows = [tuple(headings)] + [tuple(clean(v) for v in row) for row in data]
    return tuple(rows)


def write_table(value, target, name: Optional[str] = None, style: Optional[str] = "TableStyleMedium2",
                index: bool = False, auto_fit: bool = True, show_totals: bool = False):
    """
        Writes a DataFrame (pandas or polars), dict of columns or 2-d array with a
        heading row to the sheet as an Excel Table. If a table already starts at
        `target`, or a table called `name` exists, it is resized and overwritten.
        Must be called from a macro, command or event handler.

        Parameters
        ----------
        value:
            The data to write
        target:
            The top-left cell: an address string ("Sheet1!B2"), `xloil.Range` or COM Range
        name: str
            Table name. Excel requires names to be unique in a workbook
        style: str
            An Excel table style name, e.g. "TableStyleLight9"; None for Excel's default
        index: bool
            For pandas frames, whether to write the index as columns
        auto_fit: bool
            Auto-fit the column widths
        show_totals: bool
            Show the table's totals row

        Returns the COM ListObject.
    """
    headings, data = _frame_to_values(value, index)
    if data.shape[0] == 0:
        # An Excel table always has at least one (possibly empty) data row
        data = np.full((1, len(headings)), None, dtype=object)
    nrows, ncols = data.shape[0] + 1, len(headings)

    if isinstance(target, CellLocation):
        anchor = target.com_range()
    elif isinstance(target, str):
        anchor = xloil_core.app().to_com().Range(target)
    else:
        to_com = getattr(target, "to_com", None)
        anchor = to_com() if to_com else target
    anchor = anchor.Cells(1, 1)
    sheet = anchor.Worksheet

    # Find the existing table: by name or the one at the anchor
    existing = None
    if name:
        try:
            existing = sheet.ListObjects(name)
        except Exception:
            existing = None
    if existing is None:
        existing = anchor.ListObject
        if existing is not None and existing.Range.Cells(1, 1).Address != anchor.Address:
            existing = None

    new_range = com_resize(anchor, nrows, ncols)
    values = _to_com_values(headings, data)

    if existing is not None:
        # Clear old contents and resize to the new shape (Resize requires the
        # header row to stay in place)
        if existing.ShowTotals:
            existing.ShowTotals = False
        old_rows = existing.Range.Rows.Count
        old_cols = existing.Range.Columns.Count
        existing.Resize(new_range)
        # Clear cells the old table covered but the new one doesn't
        if old_rows > nrows:
            com_resize(com_offset(anchor, nrows, 0), old_rows - nrows, max(old_cols, ncols)).ClearContents()
        if old_cols > ncols:
            com_resize(com_offset(anchor, 0, ncols), old_rows, old_cols - ncols).ClearContents()
        # Write headers and body separately: assigning one array over the whole
        # range of an existing table (headers included) makes Excel drop the table
        com_resize(anchor, 1, ncols).Value = (values[0],)
        if len(values) > 1:
            com_resize(com_offset(anchor, 1, 0), len(values) - 1, ncols).Value = values[1:]
        lo = existing
    else:
        # Write the data first, then convert the range into a table: creating 
        # the table and then overwriting its whole range invalidates it. Named
        # arguments because a positional None for LinkSource is mis-marshalled.
        new_range.Value = values
        sheet.ListObjects.Add(SourceType=_XL_SRC_RANGE, Source=new_range, 
                              XlListObjectHasHeaders=_XL_YES)
        lo = anchor.ListObject

    if name and lo.Name != name:
        try:
            lo.Name = name
        except Exception as e:
            log.warn(f"write_table: could not name table '{name}': {e}")
    if style is not None:
        lo.TableStyle = style
    lo.ShowTotals = show_totals
    if auto_fit:
        lo.Range.Columns.AutoFit()
    return lo


@returner(target=object)
class TableReturn:
    """
        Return annotation: writes the returned DataFrame as an Excel Table,
        placed relative to the calling cell, and returns a short label to the
        calling cell.

        Parameters
        ----------
        name: str
            Table name (default: derived from the calling cell)
        style: str
            Excel table style name
        offset: (int, int)
            (rows, columns) offset of the table's top-left from the calling cell.
            Default (1, 0): the row below the formula
        index: bool
            Write a pandas index as columns
        label: str
            Text returned to the calling cell. "{name}" and "{rows}" are substituted
    """
    def __init__(self, name: str = None, style: str = "TableStyleMedium2",
                 offset=(1, 0), index: bool = False, label: str = "[{name}: {rows} rows]"):
        self._name = name
        self._style = style
        self._offset = offset
        self._index = index
        self._label = label

    def write(self, value):
        if xloil_core.in_wizard():
            return "[table]"
        location = caller_location()
        if location is None:
            raise CannotConvert("TableReturn can only be used from a worksheet function")

        name = self._name or ("xlo_" + "".join(ch for ch in f"{location.sheet}_{location.address}"
                                                if ch.isalnum() or ch == "_"))
        headings, data = _frame_to_values(value, self._index)
        nrows = data.shape[0]

        def do_write():
            anchor = com_offset(location.com_range(), self._offset[0], self._offset[1])
            write_table(value, anchor, name=name, style=self._style, index=self._index)

        run_after_calc(do_write)
        return self._label.format(name=name, rows=nrows)
