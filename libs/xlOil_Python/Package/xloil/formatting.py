"""
    Cell formatting applied from Python functions.

    A *formatter* describes how the cells occupied by a worksheet function's result
    should look: number format, font, fill colour, alignment, borders, column
    widths. Pass one to the `xloil.func` decorator::

        import xloil as xlo

        @xlo.func(formatter=xlo.Formatter(number_format="0.00%", bold=True))
        def growth(x: float, y: float):
            return y / x - 1

        @xlo.func(formatter=xlo.DataFrameFormatter())
        def prices(n: int):
            return make_dataframe(n)

    Excel does not allow formats to change during calculation, so xlOil records the
    calling cell and applies the formatter just after the calculation cycle.
    The formatter can also be applied directly from a macro or command with
    `Formatter.apply`.
"""

from typing import Optional, Dict, Union, Tuple, Any
from .logging import log
from ._deferred import caller_location, run_after_calc, result_shape, CellLocation

ColorSpec = Union[str, Tuple[int, int, int], int, None]

# Excel constants (avoid the win32com dependency)
_XL_ALIGN = {
    'left': -4131, 'center': -4108, 'centre': -4108, 'right': -4152,
    'general': 1, 'fill': 5, 'justify': -4130
}
_XL_VALIGN = {'top': -4160, 'center': -4108, 'centre': -4108, 'bottom': -4107}
_XL_CONTINUOUS = 1
_XL_THIN = 2
_XL_EDGES = (7, 8, 9, 10, 11, 12) # xlEdgeLeft/Top/Bottom/Right/InsideVertical/InsideHorizontal

_NAMED_COLORS = {
    'black': (0, 0, 0), 'white': (255, 255, 255), 'red': (255, 0, 0),
    'green': (0, 128, 0), 'blue': (0, 0, 255), 'yellow': (255, 255, 0),
    'grey': (128, 128, 128), 'gray': (128, 128, 128), 'lightgrey': (217, 217, 217),
    'lightgray': (217, 217, 217), 'orange': (255, 165, 0), 'navy': (0, 0, 128),
    'lightblue': (221, 235, 247), 'darkblue': (31, 78, 121),
}

def to_excel_color(color: ColorSpec) -> Optional[int]:
    """
        Converts a colour to Excel's BGR integer representation. Accepts
        '#RRGGBB', 'RRGGBB', a colour name, an (R, G, B) tuple or an int
        (passed through unchanged).
    """
    if color is None:
        return None
    if isinstance(color, int):
        return color
    if isinstance(color, str):
        name = color.strip().lower()
        if name in _NAMED_COLORS:
            r, g, b = _NAMED_COLORS[name]
        else:
            hexstr = name.lstrip('#')
            if len(hexstr) != 6:
                raise ValueError(f"Unrecognised colour '{color}'")
            r, g, b = (int(hexstr[i:i+2], 16) for i in (0, 2, 4))
    else:
        r, g, b = color
    return int(r) + 256 * int(g) + 65536 * int(b)


class Formatter:
    """
        Describes formatting to apply to a range of cells. Every parameter is
        optional: only the specified attributes are changed.

        Parameters
        ----------
        number_format: str
            An Excel number format, e.g. "0.00", "yyyy-mm-dd", "#,##0"
        bold, italic, underline: bool
        font_name: str
        font_size: float
        font_color: colour
            '#RRGGBB', a colour name, an (R,G,B) tuple or an Excel colour int
        interior_color: colour
            The cell fill colour
        horizontal_alignment: str
            'left', 'center', 'right', 'general', 'fill', 'justify'
        vertical_alignment: str
            'top', 'center', 'bottom'
        wrap_text: bool
        borders: bool or colour
            If True, draws thin borders around and between all cells. If a
            colour, uses that colour for the borders.
        auto_fit: bool
            If True, auto-fits the width of the columns occupied
        clear: bool
            If True, clears existing formats in the target range first
    """

    def __init__(self,
                 number_format: str = None,
                 bold: bool = None,
                 italic: bool = None,
                 underline: bool = None,
                 font_name: str = None,
                 font_size: float = None,
                 font_color: ColorSpec = None,
                 interior_color: ColorSpec = None,
                 horizontal_alignment: str = None,
                 vertical_alignment: str = None,
                 wrap_text: bool = None,
                 borders: Union[bool, ColorSpec] = None,
                 auto_fit: bool = False,
                 clear: bool = False):
        self.number_format = number_format
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.font_name = font_name
        self.font_size = font_size
        self.font_color = to_excel_color(font_color)
        self.interior_color = to_excel_color(interior_color)
        if horizontal_alignment is not None and horizontal_alignment not in _XL_ALIGN:
            raise ValueError(f"horizontal_alignment must be one of {list(_XL_ALIGN)}")
        if vertical_alignment is not None and vertical_alignment not in _XL_VALIGN:
            raise ValueError(f"vertical_alignment must be one of {list(_XL_VALIGN)}")
        self.horizontal_alignment = horizontal_alignment
        self.vertical_alignment = vertical_alignment
        self.wrap_text = wrap_text
        self.borders = borders
        self.auto_fit = auto_fit
        self.clear = clear

    def __repr__(self):
        attrs = ", ".join(f"{k}={v!r}" for k, v in vars(self).items()
                          if v is not None and v is not False)
        return f"{type(self).__name__}({attrs})"

    def merged(self, other: "Formatter") -> "Formatter":
        """
            Returns a new Formatter with attributes of `other` overriding this one
            where they are set.
        """
        result = Formatter()
        for k, v in vars(self).items():
            setattr(result, k, v)
        for k, v in vars(other).items():
            if v is not None and v is not False:
                setattr(result, k, v)
        return result

    def apply_to_com(self, rng):
        """ Applies this format to a COM Range object (main thread only) """
        if self.clear:
            rng.ClearFormats()
        if self.number_format is not None:
            rng.NumberFormat = self.number_format
        font = rng.Font
        if self.bold is not None: font.Bold = self.bold
        if self.italic is not None: font.Italic = self.italic
        if self.underline is not None: font.Underline = 2 if self.underline else -4142
        if self.font_name is not None: font.Name = self.font_name
        if self.font_size is not None: font.Size = self.font_size
        if self.font_color is not None: font.Color = self.font_color
        if self.interior_color is not None: rng.Interior.Color = self.interior_color
        if self.horizontal_alignment is not None:
            rng.HorizontalAlignment = _XL_ALIGN[self.horizontal_alignment]
        if self.vertical_alignment is not None:
            rng.VerticalAlignment = _XL_VALIGN[self.vertical_alignment]
        if self.wrap_text is not None:
            rng.WrapText = self.wrap_text
        if self.borders:
            color = None if self.borders is True else to_excel_color(self.borders)
            for edge in _XL_EDGES:
                try:
                    border = rng.Borders(edge)
                    border.LineStyle = _XL_CONTINUOUS
                    border.Weight = _XL_THIN
                    if color is not None:
                        border.Color = color
                except Exception:
                    pass # Inside edges fail on single cells
        if self.auto_fit:
            rng.Columns.AutoFit()

    def apply(self, target, value: Any = None):
        """
            Applies this format immediately. `target` can be an `xloil.Range`, a
            COM Range or an address string understood by ``Application.Range``.
            Must be called from a macro, command or event handler, not during
            calculation. `value` is passed through for formatters which depend on
            the value (e.g. `DataFrameFormatter`).
        """
        self.apply_value(_as_com_range(target), value)

    def apply_value(self, rng, value):
        """
            Applies the format given the COM range and the value written there.
            Subclasses override this to format parts of the range differently.
        """
        self.apply_to_com(rng)


class DataFrameFormatter(Formatter):
    """
        Formats a pandas or polars DataFrame returned by a function: the column
        headings, the index and the values can each be styled, and individual
        columns can be given their own formatter.

        Parameters
        ----------
        header: Formatter
            Applied to the heading row(s). Default: bold with a light fill and
            a bottom border
        index: Formatter
            Applied to the index column(s). Default: bold
        values: Formatter
            Applied to the data cells
        columns: dict
            Maps column name to a Formatter, applied to that column's values
        alternate_row_color: colour
            If given, fills every second data row with this colour
        auto_fit: bool
            Auto-fit all columns afterwards (default True)
        index_included: bool
            Whether the index was written (it is by default for pandas, never for
            polars). Should match the return converter options, e.g. set False if
            using ``PDFrame(index=False)``
    """

    def __init__(self,
                 header: Formatter = None,
                 index: Formatter = None,
                 values: Formatter = None,
                 columns: Dict[Any, Formatter] = None,
                 alternate_row_color: ColorSpec = None,
                 auto_fit: bool = True,
                 index_included: bool = None,
                 **kwargs):
        super().__init__(auto_fit=auto_fit, **kwargs)
        self.header = header if header is not None else \
            Formatter(bold=True, interior_color='lightblue', borders=True)
        self.index = index if index is not None else Formatter(bold=True)
        self.values = values
        self.columns = columns or {}
        self.alternate_row_color = to_excel_color(alternate_row_color)
        self.index_included = index_included

    def apply_value(self, rng, value):
        n_head, n_index, col_names, nrows = _frame_layout(value, self.index_included)
        if col_names is None:
            # Not a frame: fallback to plain formatting
            if self.values is not None:
                self.values.apply_to_com(rng)
            return

        ncols = len(col_names)
        total_rows, total_cols = nrows + n_head, ncols + n_index
        whole = rng.Resize(total_rows, total_cols)

        # Base attributes on the whole block (e.g. font_name)
        super().apply_to_com(whole)

        if n_head > 0 and self.header is not None:
            self.header.apply_to_com(whole.Resize(n_head, total_cols))
        if n_index > 0 and self.index is not None and nrows > 0:
            self.index.apply_to_com(whole.Offset(n_head, 0).Resize(nrows, n_index))
        if nrows > 0 and ncols > 0:
            body = whole.Offset(n_head, n_index).Resize(nrows, ncols)
            if self.values is not None:
                self.values.apply_to_com(body)
            if self.alternate_row_color is not None:
                for i in range(1, nrows, 2):
                    body.Rows(i + 1).Interior.Color = self.alternate_row_color
            for j, name in enumerate(col_names):
                fmt = self.columns.get(name)
                if fmt is not None:
                    fmt.apply_to_com(body.Columns(j + 1))

        if self.auto_fit:
            whole.Columns.AutoFit()


def _frame_layout(value, index_included):
    """ Returns (n_heading_rows, n_index_cols, column_names, n_data_rows) """
    try:
        import pandas as pd
        if isinstance(value, pd.Series):
            value = value.to_frame()
        if isinstance(value, pd.DataFrame):
            n_index = value.index.nlevels if index_included is not False else 0
            return value.columns.nlevels, n_index, list(value.columns), value.shape[0]
    except ImportError:
        pass
    try:
        import polars as pl
        if isinstance(value, pl.Series):
            value = value.to_frame()
        if isinstance(value, pl.DataFrame):
            return 1, 0, list(value.columns), value.height
    except ImportError:
        pass
    return 0, 0, None, 0


def DateFormatter(fmt: str = "yyyy-mm-dd", **kwargs) -> Formatter:
    """ A Formatter which applies a date number format """
    return Formatter(number_format=fmt, **kwargs)


def NumberFormatter(decimals: int = 2, thousands: bool = True, percent: bool = False,
                    **kwargs) -> Formatter:
    """ A Formatter for numbers with the given number of decimal places """
    fmt = ("#,##0" if thousands else "0") + ("." + "0" * decimals if decimals > 0 else "")
    if percent:
        fmt += "%"
    return Formatter(number_format=fmt, **kwargs)


class ConditionalFormatter(Formatter):
    """
        Applies a different formatter to each cell depending on its value.
        `rules` is a list of (predicate, Formatter) pairs: the first predicate
        returning True for a cell's value selects the formatter. Values are read
        back from the sheet after calculation.

        Example::

            red_if_negative = xlo.ConditionalFormatter([
                (lambda v: isinstance(v, (int, float)) and v < 0, xlo.Formatter(font_color='red')),
            ])

        This is evaluated cell by cell, so is best suited to small ranges. For
        large ranges use Excel's native conditional formatting.
    """
    def __init__(self, rules, default: Formatter = None):
        super().__init__()
        self.rules = list(rules)
        self.default = default

    def apply_value(self, rng, value):
        for cell in rng.Cells:
            v = cell.Value2
            for predicate, fmt in self.rules:
                try:
                    matched = predicate(v)
                except Exception:
                    matched = False
                if matched:
                    fmt.apply_to_com(cell)
                    break
            else:
                if self.default is not None:
                    self.default.apply_to_com(cell)


def _as_com_range(target):
    if isinstance(target, CellLocation):
        return target.com_range()
    if isinstance(target, str):
        import xloil_core
        return xloil_core.app().to_com().Range(target)
    to_com = getattr(target, "to_com", None)
    if to_com is not None:
        return to_com()
    return target # Assume COM range


def _schedule_format(formatter: Formatter, value):
    """
        Called from inside a worksheet function: records the caller and schedules
        `formatter` to be applied to the cells the result occupies.
    """
    import xloil_core
    try:
        if xloil_core.in_wizard():
            return
    except Exception:
        pass

    location = caller_location()
    if location is None:
        return

    rows, cols = result_shape(value,
        index=getattr(formatter, "index_included", None))

    def do_format():
        rng = location.com_range()
        if not isinstance(formatter, DataFrameFormatter):
            rng = rng.Resize(rows, cols)
        formatter.apply_value(rng, value)

    run_after_calc(do_format)
