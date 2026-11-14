# Fork changelog

Changes in [Esmaeelpour/xloil](https://github.com/Esmaeelpour/xloil) relative to
upstream [cunnane/xloil](https://github.com/cunnane/xloil). Full documentation is in
`docs/source/xlOil_Python/NewFeatures.rst` and the pages it links to.

## 0.22.1.post20

- Every xlOil binary (`xlOil.dll`, `xlOil.xll`, the `xlOil_*.dll` / `.pyd` plugins) now carries Windows version
  information: product, version, company, description, copyright and original file name, shown under
  *Properties > Details*. Unlabelled DLLs are more likely to be flagged by heuristic anti-virus: Microsoft
  Defender reported `Trojan:Win32/Bearfoos.A!ml` on the post17 `xlOil.dll`, while a Defender scan of this
  build's wheels and binaries found no threats. CI fails if a shipped xlOil binary lacks the version details.
- Natively compiled again (earlier releases since post12 reused the post12 binaries).
- Docs: new *How to Use xlOil* guide, checked by a test that runs its code examples; *Anti-virus warnings*
  section in Installing (what the detection means, how to check the file, how to report it).

## 0.22.1.post19

Found by the new install smoke test (`.github/workflows/install-smoke.yml`), which runs the documented one-liner,
read from the Installing page, on clean GitHub Windows machines for every Python version, one with Excel installed:

- Fixed: on a computer without Excel the installer pip-installed pandas and matplotlib, then crashed with a
  `FileNotFoundError` traceback. Setup now checks for Excel first and says what to do; `xloil install` gives the
  same message instead of a traceback.

The smoke test runs on every published release and can be started by hand (Actions > Install smoke test).

## 0.22.1.post18

- Fixed: in a conda-initialised PowerShell the one-liner installed into conda base even with another conda env
  activated: the new `powershell` it starts loads the profile, whose conda hook runs `conda activate base`. The
  documented command is now `powershell -NoProfile -ExecutionPolicy Bypass -c "iex (irm <url>)"`. Checked with
  a fresh conda env on Python 3.14: the env was used, doctor found no problems and Excel ran Python (first
  Excel run of the cp314 wheel). A Troubleshooting row covers it.

## 0.22.1.post17

- Fixed: `xloAttr` / `xloAttrObj` returned built-in methods uncalled: `=xloAttr("abc","upper")` gave
  `builtin_function_or_method` instead of `ABC` (only Python-defined methods and functions were called). Every
  routine is now called with the given arguments; classes and plain attributes are returned as before.

## 0.22.1.post16

From running the documented one-line installer in a real terminal (conda-initialised Windows PowerShell 5.1):

- Fixed: the one-liner `powershell ... -c "irm <url> | iex"` failed there with a parse error ("The terminator
  ... is missing from the multiline comment"), 3 runs of 3. The documented command is now
  `powershell -ExecutionPolicy Bypass -c "iex (irm <url>)"`, which ran 3 of 3 (so did `irm | Out-String | iex`).
  A Troubleshooting row names the old error. A test keeps the docs on the working form, and CI parses
  `install.ps1` with Windows PowerShell 5.1.
- Fixed: with a venv activated in a conda-initialised shell (conda base is always active there) the installer
  installed into conda base. An activated venv now comes first, then the active conda env.

## 0.22.1.post15

From upgrading the real install on a laptop (conda base, packages in the user site):

- Fixed: rerunning the installer with `-Yes` (or `xloil setup --yes`) replaced the features you had turned on
  with the first-install defaults: a Qt setup lost `xloil.gui.qtpy` and gained pandas and charts. Setup now offers
  the features the settings file already has as the defaults; a first install still gets pandas and charts.
- Fixed: `xloil config set xlOil_Python.LoadModules '["a", "b"]'` failed from Windows PowerShell 5.1, which
  strips the inner double quotes it passes to a program (`[a, b]` arrived). Bare module names in brackets are
  now taken as that list; anything else still gets the "square brackets" error.

## 0.22.1.post14

From an upgrade check (post12 to post13 with the one-line installer):

- Fixed: rerunning the installer (or `xloil setup`) reset a ribbon chosen with `xloil ribbon xlam` or `off` back to
  the COM ribbon and left `xlOil.xlam` in XLSTART, so Excel loaded both. Setup now keeps the ribbon already in the
  settings file; the signed-add-ins policy still always gets xlam, and a first install still gets the COM ribbon.
- Quieter install: pip's "A new release of pip is available" notice is off in the installer and in `xloil setup`
  (Windows PowerShell showed it as a red NativeCommandError, which looked like a failure).
- CI: the stub tests also run on Python 3.14.

## 0.22.1.post13

From a clean-install check on a laptop (python.org venv and a conda env, each with its own settings folder):

- Fixed: `xloil remove` / `uninstall` deleted only the add-in and left the `xlOil.xlam` ribbon in XLSTART; its
  macros call xlOil, so Excel showed ribbon errors at every start afterwards. Both go now; a file Excel holds
  gives "close Excel" instead of a traceback; the IntelliSense add-in stays and is named.
- Fixed: `xloil clean` started `pip uninstall` without waiting (its output came after the prompt returned)
  and skipped the logs when there was no settings file; it now deletes the settings and every `xlOil*.log`
  and waits for pip.
- Quieter install: pip's "... is installed in ... which is not on PATH" warnings are off in the installer and
  in `xloil setup` (xlOil runs as `python -m xloil...`, so PATH does not matter; the warnings alarmed people).
- Docs: *Upgrading and removing* on the Installing page.

## 0.22.1.post12

- Fixed (C++, event sink): an event that returned early (a wrong argument count, an ignored VBA error) skipped
  popping xlOil's re-entrancy stack. The stale entry made every later event of that kind look re-entrant, so it was
  dropped (`SheetSelectionChange`, `SheetChange`, `SheetActivate`, `SheetCalculate` check the stack), and after 16 such
  returns the stack array was overrun. The entry is now popped on every way out, and a nesting deeper than the stack
  drops the event instead of overrunning it.
- Docs corrected: post11's note said `SheetSelectionChange` and `SheetBeforeRightClick` do not reach Python; they do
  (re-tested: the earlier test registered a lambda, which xlOil holds weakly and dropped). `SheetBeforeDoubleClick`
  was not raised by Excel for an automated double-click, to any listener; the Events page says so and how to make a
  button that works everywhere.

## 0.22.1.post11

(The v0.22.1.post10 tag's build failed on the Linux test runners: the workbook name was split with the host's
path rules; fixed with Windows path rules. post11 is that work, released.)

Lessons from a large workbook module (a bridge design workbook) run in real Excel:

- Workbook modules: after the module registers its functions, xlOil re-enters the workbook's `#NAME?` formulas that
  were calculated before the module loaded (a full recalculation, Ctrl+Alt+F9, does not clear them). No more manual
  "Fix #NAME? errors".
- `xloil.linked_workbook(module=None)`: name a module (or its file), or let it search up the call stack, so a helper
  in a shared package, called while a workbook module loads, finds the workbook (it answered only its immediate
  caller and returned None there).
- `xloil.log.warning()` and `xloil.log.critical()` work as in Python's `logging` (only `warn` / `error` existed:
  code written with logging habits raised AttributeError inside Excel, usually in an error path).
- A function decorated with `xloil.func` keeps its own `__module__` and `__qualname__` (it reported
  `xloil.register`: code that copied `__module__` from it to register another function queued that function where
  xlOil never looks, e.g. a right-click command that then said "Cannot run the macro").
- Inside Excel, `CRYPTOGRAPHY_OPENSSL_NO_LEGACY=1` is set by default: the `cryptography` package (imported by
  xlwings, requests, ...) otherwise stops with "OpenSSL 3.0's legacy provider failed to load" in Excel's process, and
  so does any workbook module that imports it. Set it to 0 in `xlOil.ini`'s Environment to keep legacy algorithms.
- `xloil.testing.excel_array(rows)`: rows as Excel hands a range to an `xlo.Array` argument (a 2-d numpy array);
  `xloil.testing.call` passes arguments unconverted, so tests with lists missed array-only failures.
- Docs: on some Excel 365 setups the `SheetBeforeDoubleClick`, `SheetBeforeRightClick` and `SheetSelectionChange`
  events do not reach Python (seen with post9: handlers registered, never called; `SheetChange` and `SheetActivate`
  work). A button made with a data-validation drop-down and `SheetChange` works everywhere (Events page).

## 0.22.1.post9

- Fixed: `xloil doctor` (and so the end of `xloil setup`) wrongly reported "Office
  policy requires signed add-ins, so the COM ribbon can't load" after the xlOil.xlam
  ribbon was set up. It searched the whole settings file, whose comments mention
  `xloil.xloil_ribbon`; it now reads only the `LoadModules` value.

## 0.22.1.post8

- `install.ps1` no longer uses GitHub's API (api.github.com), which company networks
  often block or rate-limit, so the installer stopped with `(403) Forbidden`. It now
  reads the latest release from github.com itself and downloads the wheel directly.

## 0.22.1.post7

Re-installing or upgrading no longer needs manual steps afterwards.

- `xloil setup` keeps your own modules (anything in `LoadModules` that isn't part of
  xlOil, e.g. `mytools`) when it rebuilds the list from the features you pick.
- `install.ps1` checks for a running Excel before installing and waits for you to
  close it, instead of failing part-way with `PermissionError`.

## 0.22.1.post6

- Fixed: `xloil ribbon xlam` (and so `xloil setup` on computers with the
  signed-add-ins policy) failed with `'NoneType' object has no attribute 'settings'`:
  it loaded the ribbon code, which only runs inside Excel.

## 0.22.1.post5

- Fixed: `xloil setup` crashed with `ValueError: too many values to unpack` on
  computers where the Office signed-add-ins policy is set (the case post4's `.xlam`
  ribbon is for). The same mistake in the VBA-access check is fixed too.

## 0.22.1.post4

The *xlOil Py* ribbon tab now works on computers whose Office policy requires signed
add-ins. The compiled add-in files are unchanged (still v0.22.1.post1's).

- New `xlOil.xlam` ribbon: the same tab, drawn by an ordinary workbook add-in in
  XLSTART, for computers where the policy blocks the unsigned COM add-in.
- New command `xloil ribbon xlam|com|off` switches between the two ribbons, or turns
  the tab off.
- `xloil setup` detects the policy and picks the `.xlam` ribbon. `xloil doctor` reports
  a ribbon that can't load and says how to fix it.
- `xloil install` says to close Excel when Excel is holding the add-in file, instead
  of showing a `PermissionError` traceback.

## 0.22.1.post3

Command-line and installer fixes found setting up a locked-down work laptop. The
compiled add-in files are unchanged (still v0.22.1.post1's, see post2).

- `xloil config get/set` work with settings inside the `[[xlOil_Python.Environment]]`
  blocks, e.g. `xlOil_Python.Environment.XLOIL_PYTHON_PATH`.
- `xloil config set` refuses a value that isn't a list for list settings
  (`LoadModules`, `Plugins`, `DateFormats`, `AutoReloadExcludePaths`). Before, a typo
  could leave xlOil unable to load its modules.
- New commands: `xloil config add-module <name>`, `remove-module <name>` (keeps
  modules in load order) and `add-path <folder>` (adds a folder of your own modules
  to the search path). These replace editing lists or `xlOil.ini` by hand.
- `xloil intellisense install` no longer says to enable IntelliSense when it's
  already enabled.
- The installer tries each Python it finds and, if none starts, lists the error for
  each, with a hint when `PYTHONHOME`/`PYTHONPATH` are set, instead of
  "No Python found".

Notes for locked-down computers: if Office's policy requires add-ins to be signed,
the *xlOil Py* ribbon tab can't load (worksheet functions still work), so use the
`config` commands. If VBA project access is disabled by policy, register functions
in workbook modules with `@xlo.func(local=False)`.

## 0.22.1.post2

### Fixes
- **Excel crashed at start-up** with Qt 5 (e.g. conda's PyQt5) on some graphics
  drivers (seen with an AMD `atio6axx.dll` OpenGL driver). `xloil.gui.qtpy` turned on
  Qt's OpenGL context sharing even when Qt WebEngine wasn't installed, which loaded
  the vendor OpenGL driver into Excel. It's now only turned on when WebEngine is
  installed, and on Qt 5 xlOil always uses ANGLE (Direct3D) unless `QT_OPENGL` is
  set, so the vendor OpenGL driver is never loaded into Excel.
- Choosing a conda environment in the ribbon now sets `PYTHONHOME`. Before, Python
  failed to start after the restart.
- xlOil finds its binaries when pip installed it per-user, which pip does when the
  environment isn't writable (e.g. a conda base in C:\ProgramData).

### Easier set-up
- One-line installer: `install.ps1` is published with each release. It finds Python,
  downloads the matching wheel, installs it and runs `xloil setup`.
- `xloil setup`: guided set-up that asks which features you want, pip-installs their
  libraries, installs the add-in and runs `xloil doctor`. `--features ... --yes`
  makes it non-interactive.
- Ribbon: a **Features** menu (checkboxes for pandas, polars, charts, Qt and
  interactive charts, IntelliSense, saving the cache), with the modules added in the
  right order and missing libraries greyed out with an install hint. Environment
  entries show their Python version and whether xlOil is installed. New **Doctor**,
  **Reload** and **Settings File** buttons.
- `xloil.features.intellisense` and `xloil.features.cache_persist` turn those
  features on just by being listed in LoadModules.

### Build
- CI builds only what ships, installs only Excel for the type libraries, cancels
  superseded runs and skips docs-only changes. A release tag creates a *draft*
  release from a single build; it's published after testing.
- The binaries zip no longer includes extensions for Python 3.7 and 3.8, which are
  end-of-life and had no wheels.
- When the C++ is unchanged (checked by `tools/native_hash.py`), CI reuses the
  binaries of the last verified release (`tools/native-binaries.json`) instead of
  compiling. A fresh unsigned build of the same source was falsely flagged by
  Windows Defender's machine-learning check (`Trojan:Win32/Bearfoos.A!ml`), so this
  release ships v0.22.1.post1's binaries, built from identical C++. It also cuts a
  Python-only build from about 17 minutes to a few. Run the workflow manually with
  `force_compile` to compile anyway.

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
