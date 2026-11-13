"""
    Type converters for `polars <https://pola.rs>`_ DataFrames and Series.

    Import this module to register the converters (as with `xloil.pandas`)::

        import xloil as xlo
        import xloil.polars            # registers the converters
        import polars as pl

        @xlo.func
        def summarise(df: pl.DataFrame) -> pl.DataFrame:
            return df.group_by("region").agg(pl.col("sales").sum())

        @xlo.func
        def no_headings(df: xlo.PolarsFrame(headings=False)):
            ...

    Polars frames have no index, so a returned frame occupies one heading row
    plus one row per record. Excel serial dates in the named `dates` columns are
    converted to polars Datetime; Date/Datetime columns are written as Excel dates.
"""

try:
    import polars as pl
except ImportError:
    from ._core import XLOIL_READTHEDOCS
    if XLOIL_READTHEDOCS:
        class pl:
            class DataFrame: ...
            class Series: ...
    else:
        raise

import numpy as np
from xloil import converter, ExcelArray, CannotConvert, to_datetime


def _column_to_numpy(series: "pl.Series"):
    """ Converts a polars Series to a numpy array xlOil can write to Excel """
    dtype = series.dtype
    if dtype == pl.Date:
        series = series.cast(pl.Datetime("us"))
        dtype = series.dtype
    if dtype == pl.Datetime:
        if getattr(dtype, "time_zone", None):
            # Excel is timezone-naive: write the wall-clock time in that zone
            series = series.dt.replace_time_zone(None)
        return series.to_numpy().astype("datetime64[us]")
    if dtype in (pl.Time, pl.Duration, pl.Null) or dtype == pl.Categorical:
        return np.array([None if v is None else str(v) for v in series.to_list()], dtype=object)
    arr = series.to_numpy()
    if arr.dtype.kind not in "biuf":
        # Strings, objects and anything with nulls: None is written as an empty cell
        arr = np.array(series.to_list(), dtype=object)
    return arr


@converter(pl.DataFrame, register=True)
class PolarsFrame:
    """
        Converts between Excel ranges/arrays and *polars* DataFrames.

        Parameters
        ----------
        headings: bool
            Reading: if True (default), the first row is used for column names,
            otherwise columns are named column_0, column_1, ...
            Writing: if True (default), writes the column names as the first row
        dates: list of str
            Reading: columns containing Excel serial dates to convert to Datetime
        infer: bool
            Reading: if True (default), lets polars infer a strict column type
            where all values share one type, otherwise columns are Object
    """

    def __init__(self, headings=True, dates=None, infer=True):
        self._headings = headings
        self._dates = dates or []
        self._infer = infer

    def read(self, x):
        if isinstance(x, pl.DataFrame):
            return x
        if isinstance(x, ExcelArray):
            arr = x.to_numpy(dims=2)
        elif isinstance(x, np.ndarray):
            arr = x if x.ndim == 2 else x.reshape(-1, 1)
        else:
            raise CannotConvert(f"Unsupported type: {type(x)!r}")

        n_head = 1 if self._headings else 0
        if arr.shape[0] < n_head:
            raise CannotConvert("Expected at least one heading row")

        if n_head:
            names = [str(h) for h in arr[0, :]]
        else:
            names = [f"column_{i}" for i in range(arr.shape[1])]

        columns = {}
        for j, name in enumerate(names):
            values = arr[n_head:, j]
            if name in self._dates:
                columns[name] = pl.Series(name, to_datetime(values.astype(object)))
                continue
            values = [None if v is None or (isinstance(v, float) and np.isnan(v)) else v
                      for v in values.tolist()]
            try:
                columns[name] = pl.Series(name, values, strict=self._infer)
            except Exception:
                columns[name] = pl.Series(name, values, dtype=pl.Object)
        return pl.DataFrame(columns)

    def write(self, frame):
        if isinstance(frame, pl.LazyFrame):
            frame = frame.collect()
        if isinstance(frame, pl.Series):
            frame = frame.to_frame()
        if not isinstance(frame, pl.DataFrame):
            return frame

        import xloil_core
        columns = [_column_to_numpy(frame.get_column(name)) for name in frame.columns]
        headings = [np.array(frame.columns, dtype=object)] if self._headings else None

        return xloil_core._table_converter(
            frame.width,
            frame.height,
            columns=columns,
            index=None,
            index_name=None,
            headings=headings,
            cache_objects=False)


@converter(pl.Series, register=True)
class PolarsSeries:
    """
        Converts a single column (or row) to a polars Series, and writes a Series
        as a column with an optional heading.
    """
    def __init__(self, name: str = "", headings=False):
        self._name = name
        self._headings = headings

    def read(self, x):
        if isinstance(x, pl.Series):
            return x
        if isinstance(x, ExcelArray):
            arr = x.to_numpy(dims=1)
        elif isinstance(x, np.ndarray):
            arr = x.ravel()
        elif isinstance(x, (list, tuple)):
            arr = np.array(x, dtype=object)
        else:
            arr = np.array([x], dtype=object)
        name = self._name
        if self._headings and len(arr) > 0:
            name, arr = str(arr[0]), arr[1:]
        values = [None if v is None or (isinstance(v, float) and np.isnan(v)) else v
                  for v in arr.tolist()]
        try:
            return pl.Series(name, values)
        except Exception:
            return pl.Series(name, values, dtype=pl.Object)

    def write(self, series):
        if not isinstance(series, pl.Series):
            return series
        arr = _column_to_numpy(series)
        if self._headings:
            arr = np.concatenate([np.array([series.name], dtype=object), arr.astype(object)])
        return arr.reshape(-1, 1)


# Also allow returning LazyFrames from functions without annotation
try:
    from .type_converters import return_converters
    return_converters.add(PolarsFrame, pl.LazyFrame)
except Exception:
    pass
