================================
RTD Generators, Menus and Macros
================================

.. contents::
    :local:

RTD generators
--------------

A worksheet function written as a generator becomes a real-time-data (RTD)
function: each ``yield`` updates the cell. ``async def`` generators already
worked in xlOil; this fork adds plain ``def`` generators too:

.. code-block:: python

    import time
    import xloil as xlo

    @xlo.func
    def countdown(seconds: int):
        for remaining in range(seconds, -1, -1):
            yield remaining
            time.sleep(1)

    @xlo.func
    def price_feed(ticker: str):
        conn = connect_to_feed(ticker)      # blocking client library
        try:
            for tick in conn:
                yield tick.price
        finally:
            conn.close()

Each ``next()`` call runs in a worker thread, so a generator that blocks (on a
socket, ``time.sleep``, a database cursor) does not stall xlOil's event loop or
Excel. When the cell is deleted, its formula changes or the workbook closes,
xlOil stops the generator. Use ``try``/``finally`` to release resources. For
generators that mostly wait on I/O, ``async def`` with ``await`` scales better
because it needs no thread.

Context menus
-------------

Add entries to Excel's right-click menu:

.. code-block:: python

    @xlo.context_menu("Double values", disable_screen_updating=True)
    def double_values():
        for cell in xlo.app().Selection:
            if isinstance(cell.Value, (int, float)):
                cell.Value = cell.Value * 2

    @xlo.context_menu("Describe table", menu="List Range Popup", begin_group=True)
    def describe_table():
        df = xlo.read_table(xlo.app().Selection)
        xlo.write_table(df.describe(), "Stats!A1", name="Stats")

    @xlo.context_menu("Copy as CSV", menu=["Cell", "Row", "Column"], sub_menu="Python")
    def copy_csv(): ...

``menu``
    Which context menu(s): ``"Cell"`` (default), ``"Row"``, ``"Column"``,
    ``"Ply"`` (sheet tabs), ``"List Range Popup"`` (inside tables),
    ``"Shapes"``, or a list.
``sub_menu``
    Group items under a sub-menu.
``begin_group``, ``face_id``, ``position``
    A separator line, an Office icon id, and the 1-based position.
``name``
    The registered command name (default ``xloMenu_<function>``).
Macro options
    ``disable_screen_updating`` and the other options below.

Menu functions
--------------

``xloil.menu`` adds items to a menu shown in the **Add-ins** ribbon tab:

.. code-block:: python

    @xlo.menu("Refresh all prices", menu="Pricing")
    def refresh_prices():
        """Reloads prices from the database"""     # becomes the tooltip
        ...

    @xlo.menu("Export to CSV", menu="Pricing", sub_menu="Export")
    def export_csv(): ...

How menus work
~~~~~~~~~~~~~~

The decorated function is registered as an xlOil *command* and the menu item
runs it. Items are temporary command-bar controls: they are removed when
Excel closes and added again when the module loads. Reloading a module
replaces its items rather than duplicating them. ``xloil.menus.remove_menu_item``
and ``remove_all`` remove items, and ``menu_items()`` lists them.

For full ribbon control (tabs, galleries, ribbon-XML ``<contextMenus>``), use
:doc:`xloil.ExcelGUI <CustomGUI>`.

Macro conveniences
------------------

Commands (macros) often need to turn off screen updating and alerts, and put
the user's selection back afterwards. Pass the options to ``xloil.func`` (or
to the menu decorators):

.. code-block:: python

    @xlo.func(command=True,
              disable_screen_updating=True,
              disable_alerts=True,
              disable_calculation=True,
              restore_selection=True,
              restore_active_sheet=True)
    def rebuild_report():
        for ws in xlo.app().workbooks.active.worksheets:
            ws.activate()
            ...                    # no flicker, no dialogs, no recalcs

``disable_calculation``
    Manual calculation while running. The previous mode is restored, which
    triggers a recalculation if it was automatic.
``disable_screen_updating``
    ``Application.ScreenUpdating = False``
``disable_alerts``
    ``Application.DisplayAlerts = False``
``disable_events``
    ``Application.EnableEvents = False``
``restore_selection``
    Re-selects the previously selected cells.
``restore_active_sheet``
    Re-activates the previously active sheet.

Settings are restored even if the macro raises an exception. The same options
are available as a context manager:

.. code-block:: python

    with xlo.MacroContext(disable_screen_updating=True, restore_selection=True):
        ...

This fork also fixes ``xloil.PauseExcel``, which used to ignore its constructor
arguments.
