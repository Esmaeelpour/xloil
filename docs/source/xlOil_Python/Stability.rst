=====================
Stability and Support
=====================

Commercial Excel add-in products come with paid support and strong
backward-compatibility commitments. An open-source fork cannot match a support
contract, but it can be clear about what it promises. This page states those promises for the fork.

Compatibility
-------------

* **Upstream compatible.** The fork changes no C++ code and no existing
  function signatures. Existing xlOil add-ins, ini files and workbooks work
  unchanged. Every new behaviour is opt-in: you import a module or pass a new
  decorator argument.
* **One bug fix changes behaviour.** ``xloil.PauseExcel(...)`` now honours its
  constructor arguments. Upstream ignored them and always paused everything.
* **Versioning.** Fork releases use upstream's version plus a post-release
  number (``0.22.1.post1``), so you can see which upstream release a build
  is based on. Within a fork release series, public names documented in
  :doc:`NewFeatures` are not removed or changed incompatibly. Deprecations get
  a ``DeprecationWarning`` for at least one release before removal.
* **Platforms.** Windows x64, Excel 2013 or later (Microsoft 365 recommended),
  and Python 3.9 to 3.14 wheels. Stub mode (for tests) runs on any OS. The
  32-bit build is not produced by CI; build it locally with ``tools/stage.py``.

Quality checks
--------------

Every push runs, on GitHub Actions (free and unlimited for this public repository):

* the stub-mode unit tests (``libs/xlOil_Python/Package/tests``) on Linux and
  Windows, Python 3.9 and 3.12;
* a full Release x64 build of the add-in for every supported Python version;
* upstream's C++ unit tests;
* a smoke test that a built wheel installs and its CLI runs.

Features that need a live Excel session, such as sheet controls, menus,
tables and formatting, are checked manually against Microsoft 365 Excel before
release. The Excel-driven integration tests from upstream
(``test_PythonAutomation.py``, ``test_SpreadsheetRunner.py``) can be run
locally.

Getting help
------------

* Report bugs and requests in the fork's
  `GitHub issues <https://github.com/Esmaeelpour/xloil/issues>`_. Include
  ``xloil status`` output and the relevant part of ``%APPDATA%\xlOil\xlOil.log``
  (set ``LogLevel="debug"`` to get more detail).
* Problems in the core add-in that also happen with upstream xlOil should be
  reported to `cunnane/xloil <https://github.com/cunnane/xloil>`_ as well.
* There is no paid support or guaranteed response time. If your organisation
  needs that, budget for it explicitly: either an internal owner for this
  fork or a commercial product.

Known limitations
-----------------

* Sheet controls and interactive charts are live windows, not embedded OLE
  objects. They are not printed and need xlOil to be loaded. See
  :doc:`SheetControls`.
* Cache persistence uses pickle. Only restore from trusted workbooks.
* IntelliSense needs the separate Excel-DNA IntelliSense add-in.
* Context menus and menus use the legacy CommandBars API. It works in current
  Excel versions, but Microsoft could remove it in future.
