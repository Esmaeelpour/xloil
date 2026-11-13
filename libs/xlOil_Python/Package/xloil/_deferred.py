"""
    Internal helpers for features which must modify the workbook *after* a
    worksheet function returns (formatting, writing tables, placing controls).
    Excel does not allow a worksheet function to change anything other than its
    own return value during the calculation cycle, so these features capture the
    caller's location during the call and then schedule the work on Excel's main
    thread using `xloil.excel_callback`.
"""
from typing import NamedTuple, Optional, Callable
import xloil_core
from .logging import log


class CellLocation(NamedTuple):
    """ The workbook, sheet and local A1 address of a cell or range """
    workbook: str
    sheet: str
    address: str

    def com_range(self):
        """ Returns the COM Range object for this location (main thread only) """
        app = xloil_core.app()
        ws = app.workbooks[self.workbook].worksheets[self.sheet]
        return ws.to_com().Range(self.address)

    def worksheet(self):
        """ Returns the xloil Worksheet for this location (main thread only) """
        return xloil_core.app().workbooks[self.workbook].worksheets[self.sheet]

    def key(self) -> str:
        return f"[{self.workbook}]{self.sheet}!{self.address}"


def caller_location() -> Optional[CellLocation]:
    """
        Returns the location of the calling cell (the top-left cell for an array
        formula) or None if the function was not called from a worksheet.
        Must be called whilst the worksheet function is executing.
    """
    try:
        from .register import Caller
        caller = Caller()
        workbook = caller.workbook
        sheet = caller.sheet_name
        if workbook is None or sheet is None:
            return None
        address = caller.address(local=True)
        # Use the top-left cell of a multi-cell array formula
        address = address.split(":")[0].replace("$", "")
        return CellLocation(workbook, sheet, address)
    except Exception as e:
        log.debug(f"caller_location: could not determine caller: {e}")
        return None


def run_after_calc(func: Callable, wait: int = 0):
    """
        Schedules `func` to run on Excel's main thread once the COM API is available
        (i.e. after the current calculation cycle). Exceptions are logged rather
        than raised. Returns the future from `xloil.excel_callback`.
    """
    def logged():
        try:
            return func()
        except Exception:
            from .logging import log_except
            log_except(f"Deferred Excel callback {getattr(func, '__name__', func)} failed")

    return xloil_core.excel_callback(logged, wait=wait)


def result_shape(value, index=None, headings=None):
    """
        Estimates the (rows, columns) a value will occupy when returned to Excel
        by xlOil's default converters.  Used to size formatting and tables.
    """
    import numpy as np

    try:
        import pandas as pd
        if isinstance(value, pd.Series):
            value = value.to_frame()
        if isinstance(value, pd.DataFrame):
            n_head = value.columns.nlevels if headings is not False else 0
            n_index = value.index.nlevels if index is not False else 0
            return value.shape[0] + n_head, value.shape[1] + n_index
    except ImportError:
        pass

    try:
        import polars as pl
        if isinstance(value, pl.Series):
            return len(value) + (1 if headings is not False else 0), 1
        if isinstance(value, pl.DataFrame):
            return value.height + (1 if headings is not False else 0), value.width
    except ImportError:
        pass

    if isinstance(value, np.ndarray):
        if value.ndim == 0:
            return 1, 1
        if value.ndim == 1:
            return value.shape[0], 1
        return value.shape[0], value.shape[1]

    if isinstance(value, dict):
        return len(value), 2

    if isinstance(value, (list, tuple)):
        if len(value) == 0:
            return 1, 1
        if all(isinstance(x, (list, tuple)) for x in value):
            return len(value), max(len(x) for x in value)
        return len(value), 1

    return 1, 1
