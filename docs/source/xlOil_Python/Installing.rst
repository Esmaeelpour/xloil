====================================
Installing on Another Computer
====================================

.. contents::
    :local:

Requirements
------------

* Windows with **64-bit** desktop Excel 2013 or later (Microsoft 365 recommended).
  Check with *File > Account > About Excel*. The release wheels are 64-bit and
  cannot load into 32-bit Excel.
* A **64-bit Python, 3.9 to 3.14**, from python.org, conda (Anaconda, Miniconda,
  Miniforge) or a virtual environment built on either.
* The Microsoft Visual C++ runtime. Microsoft 365 / Click-to-Run Office includes
  it in Excel's own folder, so nothing extra is usually needed. Otherwise install
  the `VC++ Redistributable (x64) <https://aka.ms/vs/17/release/vc_redist.x64.exe>`_.
  ``xloil doctor`` tells you if it is missing.

The wheel pulls in its required packages automatically: ``numpy``, ``tomlkit`` and
``pywin32`` (used for the Excel object model).

One-line install
----------------

Open PowerShell (if you use conda or a venv, activate the environment you want
Excel to use first) and run:

.. code-block:: powershell

    powershell -ExecutionPolicy Bypass -c "irm https://github.com/Esmaeelpour/xloil/releases/latest/download/install.ps1 | iex"

The installer:

1. finds your Python (the active conda env or venv, otherwise ``py -3``, otherwise
   ``python``) and checks it is 64-bit;
2. downloads the matching wheel from the latest GitHub release and installs it;
3. runs ``xloil setup``, which asks which features you want, pip-installs the
   libraries they need, installs the Excel add-in and runs ``xloil doctor``.

Restart Excel: the *xlOil Py* ribbon tab appears. To choose a Python or skip the
questions, download ``install.ps1`` from the release and run e.g.
``.\install.ps1 -Python C:\envs\risk\python.exe -Features pandas,charts,qt -Yes``.

Install step by step
--------------------

Download the wheel for your Python version (``cp312`` = Python 3.12) from the
`GitHub release <https://github.com/Esmaeelpour/xloil/releases>`_, then from a
command prompt *in the Python environment you want Excel to use*:

.. code-block:: bat

    pip install xloil-0.22.1.post2-cp312-cp312-win_amd64.whl
    xloil setup

``xloil setup`` is the guided set-up. ``--features pandas,polars,charts,qt,intellisense,cache``
and ``--yes`` make it non-interactive. If you prefer the individual steps:
``xloil install``, then ``xloil doctor``.

If the environment isn't writable (for example a conda base installed for all users
in ``C:\ProgramData``), pip installs for your user only. xlOil finds its files
there too.

Changing settings later
-----------------------

The *xlOil Py* ribbon tab holds the common settings, so you rarely need to edit
``xlOil.ini``:

* **Features** menu: tick pandas, polars, charts, Qt widgets and interactive charts,
  IntelliSense, or saving cached objects. It adds the right modules to *Load
  Modules* in the right order. Features whose libraries aren't installed are greyed
  out, and their tooltip says what to ``pip install``.
* **Environment**: pick another Python. Entries show the Python version and whether
  xlOil is installed there. Picking a conda environment also sets ``PYTHONHOME``.
* **Doctor**, **Reload** and **Settings File** buttons: check the installation,
  reload your modules without restarting, or open ``xlOil.ini``.

Changes to the environment, and turning features off, take effect after
restarting Excel (the ribbon shows *Restart Required*).

What ``xloil install`` does
---------------------------

* copies ``xlOil.xll`` into your XLSTART folder, so Excel loads it on start;
* writes ``%APPDATA%\xlOil\xlOil.ini`` pointing at the Python that ran the
  command (keeping an existing settings file unless you pass ``--replace-ini``);
* sets ``PYTHONHOME`` automatically for conda Pythons (see below);
* removes the Windows "downloaded from the internet" mark from the add-in,
  which would otherwise make Office block it;
* warns if Excel and Python bitness differ, and if VBA project access is off
  (needed only for workbook-local functions).

Optional extras install the libraries for particular features:

.. code-block:: bat

    pip install "xloil-0.22.1.post2-cp312-cp312-win_amd64.whl[all]"      :: everything below
    pip install "xloil-...whl[pandas]"      :: DataFrames
    pip install "xloil-...whl[polars]"      :: polars
    pip install "xloil-...whl[plotting]"    :: matplotlib, plotly, Qt WebEngine charts
    pip install "xloil-...whl[qt]"          :: Qt task panes and sheet controls

Conda
-----

Conda Pythons work, including ``base``, named environments and any Python
version, as long as you install and run ``xloil`` *inside the activated
environment*:

.. code-block:: bat

    conda create -n xl python=3.12 pandas
    conda activate xl
    pip install xloil-0.22.1.post2-cp312-cp312-win_amd64.whl
    xloil install

Unlike python.org installs, an embedded conda interpreter cannot find its
standard library from ``python.exe`` alone. Without help, Python fails to start
in Excel with *"failed to get the Python codec of the filesystem encoding"*.
``xloil install``, ``xloil create`` and ``xloil config python`` detect conda (an
installation with a ``conda-meta`` folder) and write ``PYTHONHOME`` into the
settings, so this is automatic. If you edit ``xlOil.ini`` by hand, add
``PYTHONHOME='''<env folder>'''`` next to ``PYTHONEXECUTABLE``.

**Qt in conda**: use the pip wheels of PySide6, which include Qt WebEngine for
interactive charts:

.. code-block:: bat

    pip install qtpy PySide6

conda-forge's ``pyside6`` package works for task panes and sheet controls, but its
Windows build has no Qt WebEngine, so interactive charts fall back to pictures.

Pip's Qt links against Windows' own ICU library (``System32\icuuc.dll``). Many
conda environments contain a different ICU with the same file name in
``Library\bin``, which conda's Python puts on the DLL search path. Outside xlOil
this makes Qt fail with *"DLL load failed while importing QtWidgets: The
specified procedure could not be found"*. xlOil detects this situation when it
starts in Excel and loads the system ICU first, so it doesn't affect add-ins.
To use the same environment's Qt from a normal Python prompt, avoid installing
conda's ``icu``/``qt`` packages alongside pip's PySide6.

Virtual environments
--------------------

Create and activate the venv, ``pip install`` the wheel into it, and run
``xloil install`` from it. xlOil reads ``pyvenv.cfg`` to find the base Python. If
that base is a conda install, ``PYTHONHOME`` is set to it automatically. To
switch an existing installation to another environment:

.. code-block:: bat

    xloil config python C:\projects\risk\.venv

Several Pythons on one computer
-------------------------------

Excel loads a single xlOil and a single Python. The environment used is whichever
``%APPDATA%\xlOil\xlOil.ini`` points at, i.e. the one ``xloil install`` (or
``xloil config python``) was last run from. ``xloil status`` shows it.
Environment variables such as ``PYTHONPATH`` set for your user also apply inside
Excel.

Distributing an add-in to colleagues
------------------------------------

* **They have Python**: give them the wheel plus your module. They run the three
  commands above, then ``xloil config set xlOil_Python.LoadModules
  "['xloil.xloil_ribbon', 'yourmodule']"``, or point ``StartupScript`` at a
  script that sets things up.
* **Self-contained add-in file**: ``xloil create MyAddin.xll`` makes
  ``MyAddin.xll`` + ``MyAddin.ini``, configured for the current Python. Put
  ``MyAddin.py`` next to them and it loads automatically. On a computer without
  ``xloil install``, ``create`` also configures the settings file shipped with the
  binaries, so the add-in works on its own.
* **No Python on the target**: ``xloil package`` builds a PyInstaller bundle
  including Python. See :doc:`DistributingAddins`.
* **Locked-down / offline computers**: download the wheels for your dependencies
  on a connected machine (``pip download xloil-...whl -d wheels``) and install
  with ``pip install --no-index --find-links wheels xloil``.

Troubleshooting
---------------

Run ``xloil doctor`` first. It checks Excel and Python bitness, the VC++
runtime, the binaries for your Python version, required packages, the installed
add-in, internet-download blocking, the configured interpreter and conda
``PYTHONHOME``, and it says how to fix each problem it finds.

=====================================================  =================================================
Symptom                                                Fix
=====================================================  =================================================
No xlOil tab, nothing in the log                       Excel blocked the add-in: *File > Options >
                                                       Add-ins > Manage: Disabled Items*; re-run
                                                       ``xloil install`` (it also unblocks downloads)
"failed to get the Python codec of the                 Conda without PYTHONHOME: run
filesystem encoding" in ``xlOil.log``                  ``xloil config python <env>\python.exe``
``#NAME?`` for your functions                          Module not loaded: check ``LoadModules`` and the
                                                       log at ``%APPDATA%\xlOil\xlOil.log``
"DLL load failed ... QtWidgets" in a Python prompt     Conda ICU vs pip Qt clash (handled inside Excel);
                                                       see *Qt in conda* above
Interactive charts show as pictures                    No Qt WebEngine: ``pip install PySide6`` and load
                                                       ``xloil.gui.qtpy`` before ``xloil.plotting``
"is not a valid Win32 application"                     32-bit Excel or Python: use 64-bit for both
Functions using ``xloil.app()`` fail                   ``pip install pywin32``
=====================================================  =================================================

For more detail, set ``LogLevel="debug"`` with
``xloil config set Addin.LogLevel "'debug'"`` and restart Excel.
