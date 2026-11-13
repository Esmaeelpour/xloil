"""
    Optional features which switch themselves on when imported, so they can be
    enabled just by listing them in ``LoadModules`` (or ticking them in the
    ribbon's Features menu). See `FEATURES` for the catalogue used by the ribbon
    and ``xloil setup``.
"""
from typing import List, NamedTuple, Tuple


class Feature(NamedTuple):
    key: str            # short name used by `xloil setup --features`
    label: str          # shown in the ribbon / setup wizard
    modules: Tuple[str, ...]   # LoadModules entries, in load order
    requires: Tuple[str, ...]  # importable package names needed
    pip: str            # what to pip install if missing
    description: str


FEATURES: List[Feature] = [
    Feature("pandas", "pandas DataFrames", ("xloil.pandas",), ("pandas",), "pandas",
            "Pass and return pandas DataFrames"),
    Feature("polars", "polars DataFrames", ("xloil.polars",), ("polars",), "polars",
            "Pass and return polars DataFrames"),
    Feature("charts", "Charts (pictures)", ("xloil.matplotlib", "xloil.plotting"), ("matplotlib",),
            "matplotlib", "Return matplotlib, pandas and seaborn plots as pictures"),
    Feature("qt", "Qt widgets and interactive charts", ("xloil.gui.qtpy", "xloil.plotting"),
            ("qtpy",), "qtpy PySide6",
            "Qt task panes, widgets on the sheet, and interactive plotly/bokeh/altair charts"),
    Feature("intellisense", "IntelliSense tooltips", ("xloil.features.intellisense",), (), "",
            "Function help while typing formulas (needs the Excel-DNA IntelliSense add-in)"),
    Feature("cache", "Save cached objects in workbooks", ("xloil.features.cache_persist",), (), "",
            "Objects returned by @func(persist=True) are saved with the workbook"),
]

_ORDER = ["xloil.xloil_ribbon", "xloil.gui.qtpy", "xloil.pandas", "xloil.polars",
          "xloil.matplotlib", "xloil.plotting", "xloil.features.intellisense",
          "xloil.features.cache_persist"]


def feature_available(feature: Feature) -> bool:
    import importlib.util
    for package in feature.requires:
        try:
            if importlib.util.find_spec(package) is None:
                return False
        except (ImportError, ValueError):
            return False
    return True


def feature_enabled(feature: Feature, load_modules) -> bool:
    return all(m in load_modules for m in feature.modules)


def set_feature(load_modules, feature: Feature, enabled: bool) -> list:
    """
        Returns a new LoadModules list with the feature's modules added or
        removed. Modules still needed by another enabled feature are kept, and
        known modules are ordered so dependencies load first (e.g. Qt before
        plotting).
    """
    modules = list(load_modules)
    if enabled:
        for m in feature.modules:
            if m not in modules:
                modules.append(m)
    else:
        still_needed = set()
        for other in FEATURES:
            if other.key != feature.key and feature_enabled(other, modules):
                still_needed.update(other.modules)
        modules = [m for m in modules if m not in feature.modules or m in still_needed]
    return order_modules(modules)


def order_modules(modules) -> list:
    """ Puts known xlOil modules in dependency order, keeping user modules after them """
    seen = []
    for m in modules:
        if m not in seen:
            seen.append(m)
    known = [m for m in _ORDER if m in seen]
    others = [m for m in seen if m not in _ORDER]
    return known + others


def find_feature(key: str) -> Feature:
    for f in FEATURES:
        if f.key == key:
            return f
    raise KeyError(f"Unknown feature '{key}'. Choose from: {', '.join(f.key for f in FEATURES)}")
