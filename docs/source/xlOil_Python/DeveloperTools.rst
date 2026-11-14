===================================================
Developer Tools: IntelliSense, Reload, CLI, Testing
===================================================

.. contents::
    :local:

IntelliSense tooltips
---------------------

Excel shows argument help for add-in functions only in the *Insert Function*
dialog. The open-source `Excel-DNA IntelliSense
<https://github.com/Excel-DNA/IntelliSense>`_ add-in shows it in the
tooltips you see while typing a formula. ``xloil.intellisense`` feeds it from
the help text you already give xlOil (docstrings, ``help=``, ``args=``).

1. Install the IntelliSense add-in into Excel's XLSTART folder. This downloads
   the official release from GitHub:

   .. code-block:: bat

       xloil intellisense install

2. Enable description files. Add ``xloil.intellisense`` to ``LoadModules`` or
   put this in a startup script:

   .. code-block:: python

       import xloil.intellisense
       xloil.intellisense.enable()

3. Restart Excel once.

``enable()`` writes ``<AddinName>.intellisense.xml`` next to each xlOil add-in,
the file Excel-DNA IntelliSense reads. It rewrites the file whenever functions
are registered or reloaded. The description shown is the first paragraph of
the help text; argument descriptions come from ``args=`` or ``xloil.Arg``.
IntelliSense reads the files when it starts or when an add-in loads, so new
functions appear after the add-in next loads. The ``xloIntelliSenseRefresh``
command rewrites the files on demand. ``xloil intellisense remove`` uninstalls
the add-in.

Reloading without restarting Excel
----------------------------------

xlOil already watches every module it loads. Saving a module reloads it and
re-registers its functions automatically, and workbook modules reload when
the ``.py`` file changes. This fork adds explicit control for changes the
watcher can't see, such as an edited dependency or a config file:

* ``xloil.reload()`` reloads every module xlOil registered functions from.
  ``xloil.reload("mymodule")`` reloads one.
* The ``xloReload`` command does the same from the Macros dialog, a button or a menu.
* The ``=xloReloadModule("mymodule")`` worksheet function reloads one module.

Command-line installer
----------------------

See also :doc:`Installing` for installing on other computers, conda and venvs.

After ``pip install`` of the wheel:

.. code-block:: bat

    xloil install                               :: copy add-in to XLSTART, write %APPDATA%\xlOil\xlOil.ini
    xloil install --python C:\envs\risk          :: use a specific interpreter, env or venv
    xloil install --load-modules xloil.xloil_ribbon,xloil.plotting,mymod
    xloil install --startup-script C:\addins\startup.py --intellisense
    xloil status                                :: what's installed and how it's configured
    xloil doctor                                :: check for problems that stop xlOil loading
    xloil setup                                 :: guided set-up: features, libraries, add-in, checks
    xloil setup --features pandas,charts --yes  :: the same, without questions
    xloil config list
    xloil config get xlOil_Python.LoadModules
    xloil config set xlOil_Python.LoadModules "['xloil.xloil_ribbon', 'mymod']"
    xloil config set Addin.LogLevel "'info'"
    xloil config python C:\envs\risk\.venv       :: switch interpreter later
    xloil config add-path C:\work\xlo            :: folder holding your own modules
    xloil config add-module mytools             :: load your module (kept in load order)
    xloil config remove-module xloil.xloil_ribbon
    xloil remove                                :: remove the add-in, keep settings
    xloil clean                                 :: remove add-in, settings and package

``config set`` parses the value as TOML (quote strings) and keeps the
comments in ``xlOil.ini``. All ``config`` commands accept ``--ini PATH`` to
edit an add-in-specific settings file. ``create`` and ``package`` are unchanged
from upstream (see :doc:`DistributingAddins`).

Startup scripts and virtual environments
----------------------------------------

``StartupScript`` in an add-in's ``[xlOil_Python]`` section runs a script when
the add-in loads, before ``LoadModules`` are imported:

.. code-block:: toml

    [xlOil_Python]
    StartupScript = '''C:\addins\startup.py'''        # or a list of scripts
    LoadModules = ["xloil.xloil_ribbon", "mymodule"]

Relative paths resolve against the ini file's folder, and environment variables
and ``~`` are expanded. Each script runs once per add-in with ``runpy``. Use it
to set environment variables, extend ``sys.path``, configure logging, or turn
on features such as ``cache_persist.enable()`` and ``intellisense.enable()``.
A failing script is logged, and loading continues.

**Virtual environments**: point ``PYTHONEXECUTABLE`` at the venv's
``Scripts\python.exe``, for example with ``xloil config python C:\proj\.venv``.
xlOil's loader reads ``pyvenv.cfg`` to find the base installation, so nothing
else needs setting.

Testing without Excel
---------------------

On non-Windows platforms, or when ``XLOIL_STUBS=1`` is set, or with the
``xloil-stubs`` wheel installed, ``import xloil`` loads **stub mode**. The whole
API imports, decorators run and record what they declare, and nothing talks
to Excel. This is enough to unit test add-in code in any CI system.

.. code-block:: python

    # myaddin.py
    import xloil as xlo

    @xlo.func(help="Adds two numbers", args={'x': 'first', 'y': 'second'})
    def add(x: float, y: float = 1.0):
        return x + y

    @xlo.func
    def ticks(n: int):
        for i in range(n):
            yield i

.. code-block:: python

    # test_myaddin.py
    import xloil.testing as xt
    import myaddin

    def test_add():
        assert xt.call(myaddin.add, 2) == 3.0

    def test_declaration():
        spec = xt.spec(myaddin.add)
        assert spec.name == "add"
        assert [a.name for a in spec.args] == ["x", "y"]
        assert spec.args[0].help == "first"

    def test_rtd():
        assert xt.call(myaddin.ticks, 3) == [0, 1, 2]
        assert xt.take(myaddin.ticks, 2, 1000) == [0, 1]

    def test_caller():
        with xt.caller(workbook="Book1.xlsx", sheet="Data", address="C3"):
            ...

.. code-block:: bash

    pip install xloil_stubs-0.22.1.post14-py3-none-any.whl numpy pytest
    pytest

``xloil.testing``
~~~~~~~~~~~~~~~~~

``call(fn, *args)``
    Calls the Python function behind an ``@func``. Coroutines run to completion
    and generators are collected into a list.
``take(fn, n, *args)``
    The first ``n`` values from a (possibly endless) generator function.
``spec(fn)``
    The recorded declaration: ``name``, ``args``, ``help``, ``features``,
    ``is_command``, ``is_rtd``.
``caller(workbook, sheet, address)``
    Context manager setting what ``xloil.Caller()`` returns.
``fire_event(name, *args)``
    Calls handlers registered on ``xloil.event.<name>``.
``registered()``, ``pending(module)``
    Functions registered or queued for registration.
``reset()``
    Clears the in-memory cache and recorded registrations.

In stub mode ``xloil.cache`` is a working in-memory cache, ``excel_callback``
runs its callback immediately, and ``xloil.log`` writes to Python's
``logging`` (logger ``"xloil"``). Excel object model calls (``xloil.app()``)
return ``None``, so code which needs Excel should be tested against a fake or
with the integration tests in ``libs/xlOil_Python/Package/test_*.py``.

Don't install ``xloil-stubs`` in the same environment as the full ``xloil``
package: both provide the ``xloil`` module.
