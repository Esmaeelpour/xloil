# Fork changelog

Changes in [Esmaeelpour/xloil](https://github.com/Esmaeelpour/xloil) relative to
upstream [cunnane/xloil](https://github.com/cunnane/xloil). Full documentation is in
`docs/source/xlOil_Python/NewFeatures.rst` and the pages it links to.

## 0.22.1.post1

Based on upstream xlOil 0.22.1. The C++ add-in is unchanged. Everything below is in
the `xloil` Python package and is opt-in.

### Charts and sheet UI
- **Interactive charts on the sheet** (`xloil.plotting`): return plotly, bokeh or
  altair charts from worksheet functions. They show in a Qt WebEngine view placed
  next to the calling cell and update in place on recalculation. If WebEngine is
  missing they fall back to pictures.
- **More plotting libraries**: matplotlib `Axes`, pandas `df.plot()`, seaborn
  axes and grids. `xloil.plot(obj, size=, offset=, interactive=)` gives control over
  placement, and `register_plot_type` adds other libraries.
- **Sheet controls** (`xloil.SheetControl`, `place_control`): put Qt, Tk, wx or raw-HWND
  widgets directly on a worksheet. They scroll and zoom with the grid, hide on other
  sheets, and are removed with their placeholder shape.
- **Extensible GUI toolkits** (`xloil.gui.toolkits`): the `GuiToolkit` base class and
  `register_toolkit` let any toolkit's widgets be used in task panes and on sheets.
- **Context menus** (`@xloil.context_menu`) and **menus** (`@xloil.menu`, shown in the
  Add-ins tab).

### Data and functions
- **Excel Tables**: the `xloil.Table` argument type, the `xloil.TableReturn` return
  type (writes a real ListObject after calculation), and `read_table` / `write_table`.
- **Polars** (`xloil.polars`): DataFrame, Series and LazyFrame converters.
- **RTD generators**: plain `def` generators are RTD functions. Each `next()` runs in a
  worker thread.
- **Cell formatting** from functions: `@func(formatter=...)` with `Formatter`,
  `DataFrameFormatter`, `NumberFormatter`, `DateFormatter`, `ConditionalFormatter`.
- **Cache persistence** (`xloil.cache_persist`): saves cached objects inside the
  workbook and restores them under the same references on open. Use
  `@func(persist=True)`, `persist_type` or `persist_all`.
- **Macro options**: `@func(command=True, disable_screen_updating=, disable_alerts=,
  disable_calculation=, disable_events=, restore_selection=, restore_active_sheet=)` and
  the `xloil.MacroContext` context manager.
- **LRU cache** (`xloil.lru_cache`) that accepts numpy arrays, lists, dicts and
  DataFrames, with an optional TTL.
- **Workbook metadata** (`xloil.workbook_metadata`): built-in and custom document
  properties, plus a JSON store in a Custom XML Part.

### Developer experience
- **IntelliSense** (`xloil.intellisense`): writes Excel-DNA IntelliSense description
  files from function help. `xloil intellisense install` fetches the add-in.
- **Reload**: `xloil.reload()`, the `xloReload` command and the `xloReloadModule`
  worksheet function. Automatic reload on save is upstream behaviour.
- **CLI**: `xloil status`, `xloil config get/set/list/python`, and
  `xloil install --python/--load-modules/--startup-script/--intellisense/--replace-ini`.
- **Startup scripts**: the `StartupScript` setting in `[xlOil_Python]` runs before
  `LoadModules`.
- **Stub mode for testing**: `import xloil` works without Excel (on non-Windows, with
  `XLOIL_STUBS=1`, or with the new `xloil-stubs` wheel). `xloil.testing` provides
  `call`, `take`, `spec`, `caller`, `fire_event`, `registered`, `pending` and `reset`.

### Installation
- **Conda support**: `xloil install`, `xloil create` and `xloil config python` set
  `PYTHONHOME` for conda Pythons (base, envs, and venvs built on conda). Without it,
  Python failed to start inside Excel ("failed to get the Python codec of the
  filesystem encoding").
- **`xloil doctor`** checks Excel/Python bitness, the VC++ runtime, the binaries for
  this Python, required packages, the add-in install, download blocking and the
  configured interpreter, and says how to fix each problem.
- `pywin32` is now a required dependency on Windows (xlOil's Excel object model uses
  it). New extras: `all`, `pandas`, `polars`, `plotting`, `qt`, `test`.
- `xloil install` warns on Excel/Python bitness mismatch and removes the
  internet-download mark that makes Office block XLLs.
- `xloil create` also configures the settings file shipped with the binaries, so a
  created add-in works on a computer where `xloil install` was never run.
- **Pip Qt in conda**: xlOil preloads Windows' ICU when a conda environment's own
  `icuuc.dll` would otherwise break pip-installed PySide6/PyQt ("procedure could not
  be found"), so interactive charts work in conda environments.
- Clearer error when a chart can be shown neither interactively nor as a picture.
- New docs page: Installing on Another Computer.
- Tested end to end in Excel with python.org Python 3.9, a conda Python 3.12 env, and
  a venv built on conda Python 3.10.

### Fixes
- Invalid escape sequence warning in `xloil.jupyter` on Python 3.12+.
- `xloil.PauseExcel(...)` now honours its constructor arguments. Upstream always
  paused events, calculation, screen updating and alerts.

### Build and release
- A GitHub Actions workflow (`build-release.yml`) builds Release x64, wheels for
  Python 3.9 to 3.14 and the stubs wheel, runs the tests, and publishes a GitHub
  Release on `v*` tags. GitHub Actions minutes are unlimited for public repositories.
- The upstream GitLab mirror workflow now only runs in the upstream repository.
