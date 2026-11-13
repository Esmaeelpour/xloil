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

Install in three commands
-------------------------

Download the wheel for your Python version (``cp312`` = Python 3.12) from the
`GitHub release <https://github.com/Esmaeelpour/xloil/releases>`_, then from a
command prompt *in the Python environment you want Excel to use*:

.. code-block:: bat

    pip install xloil-0.22.1.post1-cp312-cp312-win_amd64.whl
    xloil install
    xloil doctor

Start Excel. The *xlOil* ribbon tab appears.

``xloil install``:

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

    pip install "xloil-0.22.1.post1-cp312-cp312-win_amd64.whl[all]"      :: everything below
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
    pip install xloil-0.22.1.post1-cp312-cp312-win_amd64.whl
    xloil install

Unlike python.org installs, an embedded conda interpreter cannot find its
standard library from ``python.exe`` alone. Without help, Python fails to start
in Excel with *"failed to get the Python codec of the filesystem encoding"*.
``xloil install``, ``xloil create`` and ``xloil config python`` detect conda (an
installation with a ``conda-meta`` folder) and write ``PYTHONHOME`` into the
settings, so this is automatic. If you edit ``xlOil.ini`` by hand, add
``PYTHONHOME='''<env folder>'''`` next to ``PYTHONEXECUTABLE``.

**Qt in conda**: install Qt from conda, not pip:

.. code-block:: bat

    conda install -c conda-forge pyside6 qtpy

A pip-installed PySide6 can clash with Qt DLLs that other conda packages put in
the environment's ``Library\bin``, which xlOil adds to ``PATH``. The symptom is
*"DLL load failed while importing QtWidgets: The specified procedure could not
be found"*.

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
"DLL load failed ... QtWidgets"                        Conda + pip Qt: ``conda install pyside6``
"is not a valid Win32 application"                     32-bit Excel or Python: use 64-bit for both
Functions using ``xloil.app()`` fail                   ``pip install pywin32``
=====================================================  =================================================

For more detail, set ``LogLevel="debug"`` with
``xloil config set Addin.LogLevel "'debug'"`` and restart Excel.
