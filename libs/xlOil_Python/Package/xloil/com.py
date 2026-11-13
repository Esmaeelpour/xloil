from . import _core

# TODO: support comtypes
class _Win32ComConstants:
    def __getattr__(self, name):
        from win32com.client import constants
        return getattr(constants, name)


constants = _Win32ComConstants()
"""
    Contains the numeric value for enums used in the Excel.Application API.
    For example:

    ::

        from xloil import constants as xlc
        xloil.app().Calculation = xlc.xlCalculationManual
        xloil.app().Selection.PasteSpecial(Paste=xlc.xlPasteFormulas)

"""

class MacroContext:
    """
    A context manager which temporarily changes Excel application settings 
    whilst a macro runs and restores them afterwards, even if an exception is
    raised. Each option defaults to False (i.e. leave that setting alone).

    The same options are available directly on the `xloil.func` decorator, e.g.
    ``@xloil.func(command=True, disable_screen_updating=True)``.

    Parameters
    ----------
    disable_calculation: bool
        Sets the calculation mode to manual (``Application.Calculation``)
    disable_screen_updating: bool
        Turns off ``Application.ScreenUpdating``
    disable_alerts: bool
        Turns off ``Application.DisplayAlerts``
    disable_events: bool
        Turns off ``Application.EnableEvents``
    restore_selection: bool
        Re-selects the previously selected cells on exit
    restore_active_sheet: bool
        Re-activates the previously active sheet on exit

    Examples
    --------

    ::

        with xloil.MacroContext(disable_screen_updating=True, restore_selection=True):
            for cell in xloil.app().range("A1:A100"):
                cell.Select()
                ...

    """
    # The below constant equals constants.xlCalculationManual but
    # avoids the dependency on win32com
    _XL_CALCULATION_MANUAL = -4135

    def __init__(self, 
                 disable_calculation=False, 
                 disable_screen_updating=False, 
                 disable_alerts=False, 
                 disable_events=False,
                 restore_selection=False,
                 restore_active_sheet=False):
        self._options = dict(
            calculation=disable_calculation,
            screen_updating=disable_screen_updating,
            alerts=disable_alerts,
            events=disable_events)
        self._restore_selection = restore_selection
        self._restore_active_sheet = restore_active_sheet
        self._saved = {}
        self._selection = None
        self._active_sheet = None

    _PROPERTIES = dict(
        calculation=('Calculation', _XL_CALCULATION_MANUAL),
        screen_updating=('ScreenUpdating', False),
        alerts=('DisplayAlerts', False),
        events=('EnableEvents', False))

    def __enter__(self):
        app = _core.app()
        self._saved = {}

        if self._restore_active_sheet:
            try:
                self._active_sheet = app.ActiveSheet
            except Exception:
                self._active_sheet = None

        if self._restore_selection:
            try:
                self._selection = app.Selection
            except Exception:
                self._selection = None

        for key, enabled in self._options.items():
            if not enabled:
                continue
            prop, value = self._PROPERTIES[key]
            self._saved[prop] = getattr(app, prop)
            setattr(app, prop, value)

        return self

    def __exit__(self, exc_type, exc_value, traceback):
        app = _core.app()
        # Restore in reverse order: calculation last so a recalc happens 
        # after screen updating/events are back on
        for prop in reversed(list(self._saved)):
            try:
                setattr(app, prop, self._saved[prop])
            except Exception as e:
                from .logging import log
                log.warn(f"MacroContext: failed to restore {prop}: {e}")

        if self._active_sheet is not None:
            try:
                self._active_sheet.Activate()
            except Exception:
                pass

        if self._selection is not None:
            try:
                # Selecting a range on a non-active sheet fails, so activate first
                self._selection.Worksheet.Activate()
                self._selection.Select()
            except Exception:
                pass

        self._saved = {}
        self._selection = None
        self._active_sheet = None
        return False


class PauseExcel(MacroContext):
    """
    A context manager which pauses Excel by disabling events, turning off screen updating 
    and switching to manual calculation mode.  Which of these changes are applied can
    be controlled by parameters to the constructor - the default is to apply all of them.
    Previous settings are restored when the context scope closes.

    Parameters
    ----------
    events: bool (default False)
        if False, pauses Excel's event model (Application.EnableEvents in VBA)

    calculation: bool (default False)
        if False, sets the calculation mode to manual (Application.Calculation in VBA)

    screen_updating: bool
        if False, disables screen updating in Excel (Application.ScreenUpdating in VBA)

    alerts: bool
        if False, disables alerts in Excel (Application.DisplayAlerts in VBA)

    """
    def __init__(self, events=False, calculation=False, screen_updating=False, alerts=False):
        # Note the inverted sense of the flags compared to MacroContext: here
        # False means "turn this off"
        super().__init__(
            disable_calculation=not calculation,
            disable_screen_updating=not screen_updating,
            disable_alerts=not alerts,
            disable_events=not events)


def fix_name_errors(workbook):
    """
        Marks #NAME! errors in current workbook for recalculation. Can be used
        when missing addin functions have caused #NAME! errors - these errors cannot
        simply be resolved by full recalc (Ctrl-Alt-F9).

        May not be performant in large workbooks with many errors.
    """
    for ws in workbook.worksheets:
        errors = ws.used_range.special_cells("formulas", "errors")
        if errors is None:
            continue
        for cell in errors:
            if cell.Value2 == _core.CellError.NAME.value:
                cell.Formula = cell.Formula
