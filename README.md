xlOil
=====

> **This is a fork of [cunnane/xloil](https://github.com/cunnane/xloil)** that adds
> features for charts, sheet widgets, data handling and developer tools. All additions are in the Python package
> and are opt-in; the compiled add-in is unchanged. See
> [Fork Features](docs/source/xlOil_Python/NewFeatures.rst) for the full list and
> [CHANGELOG-fork.md](CHANGELOG-fork.md) for what changed.
>
> | Area | Added |
> |---|---|
> | Charts | Interactive plotly / bokeh / altair charts on the sheet; matplotlib, pandas, seaborn pictures; custom libraries |
> | Sheet UI | Qt/Tk/wx widgets placed on the worksheet; pluggable GUI toolkit bridge; right-click context menus; Add-ins menus |
> | Data | Excel Tables in and out; polars; cell formatting from functions; cache objects saved in the workbook; LRU cache; workbook metadata |
> | Functions | Plain generators as RTD functions; macro options (screen updating, alerts, calc, restore selection) |
> | Dev tools | Excel-DNA IntelliSense tooltips; `xloil.reload()`; `xloil status/config` CLI; `StartupScript`; stub mode + `xloil.testing` + `xloil-stubs` wheel |
>
> Binaries and wheels: [GitHub Releases](https://github.com/Esmaeelpour/xloil/releases).
> Install (python.org, conda or venv; 64-bit Excel and Python 3.9–3.14):
> `pip install xloil-<version>-cp3XX-cp3XX-win_amd64.whl`, then `xloil install`, then `xloil doctor`.
> See [Installing on Another Computer](docs/source/xlOil_Python/Installing.rst).
> Then [How to Use xlOil](docs/source/xlOil_Python/HowToUse.rst): your first function, everyday recipes and fixes.

xlOil provides framework for interacting with Excel in different programming
languages. It gives a way to write functions in a language and have them 
appear in Excel as worksheet functions or macros or have them control GUI elements. 
For example, the Python bindings can replace VBA in almost all use cases 
and even provide functionality not available from VBA.

xlOil is designed to have very low overheads when calling your worksheet 
functions.

xlOil supports different languages via plugins. The languages currently 
supported are:

- C++
- Python
- SQL

In addition there is *xlOil_Utils* which contains some handy tools which Microsoft
never quite got around to adding.

The latest stable documentation is here: https://xloil.readthedocs.io/en/stable.

xlOil features
--------------

* Python
  - Concise syntax to declare an Excel function
  - Optional type checking of function parameters
  - Supports keyword arguments
  - Choice of globally declared functions or code modules limited to a single workbook
    like VBA workbook-level functions
  - Tight integration with *numpy* - very low overheads for array functions
  - Understands python tuples, lists, dictionarys and *pandas* dataframes
  - Async functions
  - RTD functions and on-the-fly RTD server creation
  - Macro type functions which write to ranges on the sheet
  - Access to the Excel Application object and the full Excel object model
  - Drive Excel through COM automation
  - Hook Excel events
  - Pass any python object back to Excel and back into another function
  - Simple add-in deployment
  - Two-way connection to *Jupyter* notebooks: run worksheet functions in *Jupyter* and query variables
    in the *jupyter* kernel
  - Create Ribbon toolbars and Custom Task Panes
  - Return *matplotlib* plots and images from worksheet functions
* C++
  - Safe and convenient wrappers around most things in the C-API
  - Concise syntax to declare Excel functions: registration is automatic
  - Deal with Excel variants, Ranges, Arrays and strings in a natural C++ fashion
  - Object cache allows returning opaque objects to Excel and passing them back to other functions
  - Simplified RTD server creation
  - RTD-based background calculation
  - Create Ribbon toolbars and Custom Task Panes  
* SQL
  - Create tables from Excel ranges and arrays
  - Query and join them with the full sqlite3 SQL syntax
* Utils: very fast functions to:
  - Sort on multiple columns
  - Split and join strings
  - Make arrays from blocks


Supporting other languages
--------------------------

You can use xlOil as an end-user of these plugins or you can use it to write
you own language bindings (and ideally add them to the repo).