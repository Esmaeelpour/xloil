"""
    Stub-mode unit tests for the features added on top of upstream xlOil.
"""
import asyncio
import datetime as dt
import sys
import time
import xml.etree.ElementTree as ET
from pathlib import Path

import numpy as np
import pytest

import xloil as xlo
import xloil.testing as xt
import xloil_core
import xloil._core

from conftest import FakeCom, FakeRange, FakeWorkbook


# ---------------------------------------------------------------- stub mode

def test_stub_mode_active():
    assert xt.in_stub_mode()


def test_func_records_spec():
    @xlo.func(help="Adds", args={'x': 'first'})
    def add(x: float, y: float = 2.0):
        return x + y

    spec = xt.spec(add)
    assert spec.name == "add"
    assert [a.name for a in spec.args] == ["x", "y"]
    assert spec.args[0].help == "first"
    assert "macro" in spec.features
    assert xt.call(add, 1.0) == 3.0
    assert spec in xt.pending(sys.modules[__name__])


def test_caller_context():
    @xlo.func
    def where():
        return xlo.Caller().address()

    with xt.caller("Prices.xlsx", "Data", "B4"):
        assert xt.call(where) == "[Prices.xlsx]Data!B4"


# ------------------------------------------------------------ RTD generators

def test_sync_generator_becomes_rtd():
    @xlo.func
    def ticker(n: int):
        for i in range(n):
            yield i

    spec = xt.spec(ticker)
    assert "rtd" in spec.features
    assert xt.call(ticker, 3) == [0, 1, 2]


def test_sync_generator_runs_in_executor():
    from xloil.register import _sync_generator_to_async

    def slow(n):
        for i in range(n):
            time.sleep(0.01)
            yield i * 10

    agen = _sync_generator_to_async(slow)

    async def collect():
        return [x async for x in agen(3)]

    assert asyncio.run(collect()) == [0, 10, 20]


def test_take_stops_infinite_generator():
    @xlo.func
    def forever():
        i = 0
        while True:
            yield i
            i += 1

    assert xt.take(forever, 4) == [0, 1, 2, 3]


def test_async_generator_still_supported():
    @xlo.func
    async def agen(n: int):
        for i in range(n):
            yield i

    assert "rtd" in xt.spec(agen).features
    assert xt.call(agen, 2) == [0, 1]


# ------------------------------------------------------------- macro options

def test_macro_options_wrap_function(monkeypatch):
    calls = []

    class FakeContext:
        def __init__(self, **kwargs):
            calls.append(kwargs)
        def __enter__(self): return self
        def __exit__(self, *a): return False

    import xloil.com
    monkeypatch.setattr(xloil.com, "MacroContext", FakeContext)

    @xlo.func(command=True, disable_screen_updating=True, restore_selection=True)
    def cmd():
        return "done"

    spec = xt.spec(cmd)
    assert "command" in spec.features
    assert spec.func() == "done"
    assert calls == [{'disable_screen_updating': True, 'restore_selection': True}]


def test_macro_options_rejected_for_async():
    # func() logs and returns the undecorated function on error
    @xlo.func(disable_alerts=True)
    async def bad():
        return 1
    assert not hasattr(bad, "_xloil_spec")


def test_macro_context_restores(monkeypatch):
    class App:
        Calculation = -4105
        ScreenUpdating = True
        DisplayAlerts = True
        EnableEvents = True
        ActiveSheet = None
        Selection = None

    app = App()
    monkeypatch.setattr(xloil._core, "app", lambda: app)
    with xlo.MacroContext(disable_calculation=True, disable_screen_updating=True,
                          disable_alerts=True):
        assert app.Calculation == -4135
        assert app.ScreenUpdating is False
        assert app.DisplayAlerts is False
        assert app.EnableEvents is True
    assert app.Calculation == -4105
    assert app.ScreenUpdating is True
    assert app.DisplayAlerts is True


def test_pause_excel_honours_constructor_args(monkeypatch):
    class App:
        Calculation = -4105
        ScreenUpdating = True
        DisplayAlerts = True
        EnableEvents = True
    app = App()
    monkeypatch.setattr(xloil._core, "app", lambda: app)
    # Upstream ignored these arguments; screen_updating=True means "leave it on"
    with xlo.PauseExcel(screen_updating=True):
        assert app.ScreenUpdating is True
        assert app.EnableEvents is False
    assert app.EnableEvents is True


# ---------------------------------------------------------------- LRU cache

def test_lru_cache_arrays():
    calls = []

    @xlo.lru_cache(maxsize=2)
    def f(x, opts=None):
        calls.append(1)
        return float(np.sum(x))

    a = np.arange(6.0).reshape(2, 3)
    assert f(a) == 15.0
    assert f(a.copy()) == 15.0          # equal contents -> hit
    assert len(calls) == 1
    assert f(a, opts={'k': [1, 2]}) == 15.0
    assert f(a, opts={'k': [1, 2]}) == 15.0
    assert len(calls) == 2
    info = f.cache_info()
    assert info.hits == 2 and info.misses == 2 and info.currsize == 2
    f(np.ones(3))                       # evicts oldest
    assert f.cache_info().currsize == 2
    f.cache_clear()
    assert f.cache_info().currsize == 0


def test_lru_cache_distinguishes_types():
    @xlo.lru_cache
    def f(x):
        return type(x).__name__
    assert f(1) == "int"
    assert f(1.0) == "float"


def test_lru_cache_ttl():
    calls = []
    @xlo.lru_cache(ttl=0.05)
    def f(x):
        calls.append(x)
        return x
    f(1); f(1)
    assert len(calls) == 1
    time.sleep(0.08)
    f(1)
    assert len(calls) == 2


def test_lru_cache_preserves_signature_for_func():
    @xlo.func
    @xlo.lru_cache(maxsize=8)
    def g(x: int, y: str = "a"):
        return f"{x}{y}"
    spec = xt.spec(g)
    assert [a.name for a in spec.args] == ["x", "y"]
    assert xt.call(g, 3) == "3a"


def test_lru_cache_pandas():
    pd = pytest.importorskip("pandas")
    calls = []
    @xlo.lru_cache
    def f(df):
        calls.append(1)
        return len(df)
    df = pd.DataFrame({'a': [1, 2], 'b': [3, 4]})
    f(df); f(df.copy())
    assert len(calls) == 1
    f(df.assign(b=[3, 5]))
    assert len(calls) == 2


# ---------------------------------------------------------------- formatting

def test_colors():
    from xloil.formatting import to_excel_color
    assert to_excel_color("#FF0000") == 255
    assert to_excel_color((0, 0, 255)) == 255 * 65536
    assert to_excel_color("white") == 0xFFFFFF
    assert to_excel_color(123) == 123
    with pytest.raises(ValueError):
        to_excel_color("#12")


def test_formatter_applies_attributes():
    rng = FakeRange()
    xlo.Formatter(number_format="0.00", bold=True, interior_color="#00FF00",
                  horizontal_alignment="center").apply_to_com(rng)
    assert rng.writes()["NumberFormat"] == "0.00"
    assert rng.Font.writes()["Bold"] is True
    assert rng.Interior.writes()["Color"] == 65280
    assert rng.writes()["HorizontalAlignment"] == -4108


def test_formatter_validation():
    with pytest.raises(ValueError):
        xlo.Formatter(horizontal_alignment="middle")


def test_number_formatter():
    assert xlo.NumberFormatter(2).number_format == "#,##0.00"
    assert xlo.NumberFormatter(1, thousands=False, percent=True).number_format == "0.0%"


def test_dataframe_formatter_regions():
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame({'a': [1, 2, 3], 'b': [4.0, 5.0, 6.0]})
    rng = FakeRange(1, 1)
    fmt = xlo.DataFrameFormatter(columns={'b': xlo.NumberFormatter(1)}, alternate_row_color="grey")
    fmt.apply_value(rng, df)
    ranges = [r for r in rng.log if isinstance(r, FakeRange)]
    # whole block: 1 heading row + 3 rows, 1 index col + 2 cols
    assert any(r.nrows == 4 and r.ncols == 3 for r in ranges)
    # header region
    assert any(r.row == 1 and r.nrows == 1 and r.ncols == 3 for r in ranges)
    # column 'b' body: row 2, col 3
    col_b = [r for r in ranges if r.row == 2 and r.col == 3 and r.ncols == 1]
    assert col_b and col_b[-1].writes().get("NumberFormat") == "#,##0.0"


def test_result_shape():
    from xloil._deferred import result_shape
    pd = pytest.importorskip("pandas")
    assert result_shape(np.zeros((3, 4))) == (3, 4)
    assert result_shape(np.zeros(5)) == (5, 1)
    assert result_shape([[1, 2], [3, 4, 5]]) == (2, 3)
    assert result_shape(7) == (1, 1)
    df = pd.DataFrame({'a': [1, 2]})
    assert result_shape(df) == (3, 2)
    assert result_shape(df, index=False) == (3, 1)


def test_formatter_scheduled_after_call(monkeypatch):
    applied = []

    class Recorder(xlo.Formatter):
        def apply_value(self, rng, value):
            applied.append((rng, value))

    import xloil._deferred as deferred
    monkeypatch.setattr(deferred.CellLocation, "com_range", lambda self: FakeRange())

    @xlo.func(formatter=Recorder())
    def three():
        return np.ones((3, 2))

    with xt.caller(address="C5"):
        result = xt.spec(three).func()
    assert result.shape == (3, 2)
    assert len(applied) == 1
    rng, value = applied[0]
    assert (rng.nrows, rng.ncols) == (3, 2)


def test_formatter_rejected_for_rtd():
    @xlo.func(formatter=xlo.Formatter(bold=True))
    def gen():
        yield 1
    assert not hasattr(gen, "_xloil_spec")


# ----------------------------------------------------------------- metadata

def test_store_encoding_roundtrip():
    from xloil.metadata import encode_store, decode_store, STORE_NAMESPACE
    values = {'a': b'hello', 'b<&>': b'\x00\x01\xff'}
    xml = encode_store(values)
    assert decode_store(xml) == values
    ET.fromstring(xml)  # valid XML


def test_json_store(fake_workbook):
    from xloil.metadata import _JsonStore
    store = _JsonStore(fake_workbook)
    store['run'] = {'rows': 3, 'user': 'me'}
    store['n'] = 5
    assert store['run'] == {'rows': 3, 'user': 'me'}
    assert sorted(store) == ['n', 'run']
    del store['n']
    assert list(store) == ['run']
    # Only one xml part is ever kept
    assert len(fake_workbook.CustomXMLParts.parts) == 1
    store.clear()
    assert len(fake_workbook.CustomXMLParts.parts) == 0


# ----------------------------------------------------------- cache persist

def test_cache_ref_parsing():
    from xloil.cache_persist import _ref_workbook, _ref_to_key
    ref = "〈[Book1.xlsx]Sheet1!R1C1,A"
    assert _ref_workbook(ref) == "Book1.xlsx"
    assert _ref_to_key(ref) == "Book1.xlsx]Sheet1!R1C1"


def test_cache_persist_roundtrip(fake_workbook, monkeypatch):
    import xloil.cache_persist as cp
    monkeypatch.setattr(cp, "_workbook_com", lambda wb: fake_workbook)
    cache = xloil_core.cache

    keep = cp.mark_persistent({'model': [1, 2, 3]})
    ref_keep = cache.add(keep, key="Book1.xlsx]Sheet1!R1C1")
    ref_skip = cache.add({'tmp': 1}, key="Book1.xlsx]Sheet1!R2C1")
    ref_other = cache.add(cp.mark_persistent([9]), key="Other.xlsx]Sheet1!R1C1")

    assert cp.save("Book1.xlsx") == 1

    cache.clear()
    assert cp.restore("Book1.xlsx") == 1
    assert cache.get(ref_keep) == {'model': [1, 2, 3]}
    assert ref_skip not in cache
    assert ref_other not in cache


def test_cache_persist_types(fake_workbook, monkeypatch):
    import xloil.cache_persist as cp
    monkeypatch.setattr(cp, "_workbook_com", lambda wb: fake_workbook)

    class Model:
        def __init__(self, v): self.v = v

    # Module-level class needed for pickle: register via module attribute
    Model.__qualname__ = "Model"
    globals()['Model'] = Model
    cp.persist_type(Model)
    try:
        xloil_core.cache.add(Model(4), key="Book1.xlsx]S!R1C1")
        assert cp.save("Book1.xlsx") == 1
    finally:
        cp._persist_types.discard(Model)


def test_persist_flag_marks_result():
    import xloil.cache_persist as cp
    @xlo.func(persist=True)
    def make():
        return {'big': 1}
    obj = xt.spec(make).func()
    assert id(obj) in cp._marked_ids


# ---------------------------------------------------------------- polars

def test_polars_read_and_columns():
    pl = pytest.importorskip("polars")
    from xloil.polars import PolarsFrame, _column_to_numpy
    arr = np.array([["name", "qty", "price"], ["a", 1.0, 2.5], ["b", 2.0, None]], dtype=object)
    df = PolarsFrame.__wrapped__().read(arr)
    assert df.columns == ["name", "qty", "price"]
    assert df.height == 2
    assert df["price"][1] is None

    dates = pl.Series("d", [dt.date(2024, 1, 2), None])
    out = _column_to_numpy(dates)
    assert out.dtype.kind == "M"
    tz = pl.Series("t", [dt.datetime(2024, 1, 2, 12)]).dt.replace_time_zone("UTC")
    assert _column_to_numpy(tz).dtype.kind == "M"
    strs = _column_to_numpy(pl.Series("s", ["x", None]))
    assert list(strs) == ["x", None]


def test_polars_no_headings():
    pytest.importorskip("polars")
    from xloil.polars import PolarsFrame
    df = PolarsFrame.__wrapped__(headings=False).read(np.array([[1.0, 2.0], [3.0, 4.0]]))
    assert df.columns == ["column_0", "column_1"]
    assert df.height == 2


def test_polars_series():
    pytest.importorskip("polars")
    from xloil.polars import PolarsSeries
    s = PolarsSeries.__wrapped__(headings=True).read(np.array(["h", 1.0, 2.0], dtype=object))
    assert s.name == "h" and s.to_list() == [1.0, 2.0]


# ---------------------------------------------------------------- tables

def test_frame_to_values_pandas():
    pd = pytest.importorskip("pandas")
    from xloil.tables import _frame_to_values, _to_com_values
    df = pd.DataFrame({'a': [1, None], 'when': pd.to_datetime(["2024-01-01", None])})
    heads, data = _frame_to_values(df, index=False)
    assert heads == ['a', 'when']
    assert data.shape == (2, 2)
    assert data[1, 0] is None and data[1, 1] is None
    rows = _to_com_values(heads, data)
    assert rows[0] == ('a', 'when')
    assert isinstance(rows[1][1], dt.datetime)


def test_frame_to_values_dict_and_index():
    pd = pytest.importorskip("pandas")
    from xloil.tables import _frame_to_values
    heads, data = _frame_to_values({'x': [1, 2, 3], 'y': [4]}, index=False)
    assert heads == ['x', 'y'] and data.shape == (3, 2) and data[2, 1] is None
    df = pd.DataFrame({'v': [1]}, index=pd.Index(['k'], name='key'))
    heads, _ = _frame_to_values(df, index=True)
    assert heads == ['key', 'v']


def test_frame_to_values_polars():
    pl = pytest.importorskip("polars")
    from xloil.tables import _frame_to_values
    heads, data = _frame_to_values(pl.DataFrame({'a': [1, 2], 'b': ['x', 'y']}), index=False)
    assert heads == ['a', 'b'] and data[1, 1] == 'y'


# ------------------------------------------------------------- intellisense

def test_intellisense_xml():
    from xloil.intellisense import spec_to_xml, document_xml, NAMESPACE

    @xlo.func(help="Computes things.\n\nLonger details here.", args={'a': 'The A'})
    def calc(a, b=1, *rest):
        return a

    xml = document_xml([spec_to_xml(xt.spec(calc))])
    root = ET.fromstring(xml)
    ns = {'i': NAMESPACE}
    fn = root.find("i:FunctionInfo/i:Function", ns)
    assert fn.get("Name") == "calc"
    assert fn.get("Description") == "Computes things."
    args = [a.get("Name") for a in fn.findall("i:Argument", ns)]
    assert args == ["a", "b", "rest"]
    assert fn.find("i:Argument", ns).get("Description") == "The A"


def test_intellisense_writes_file(tmp_path):
    from xloil.intellisense import write_addin_file

    @xlo.func
    def f1(x): return x

    class Addin:
        pathname = str(tmp_path / "MyAddin.xll")
        def functions(self):
            return [xt.spec(f1)]

    path = write_addin_file(Addin())
    assert path == tmp_path / "MyAddin.intellisense.xml"
    assert 'Name="f1"' in path.read_text(encoding="utf-8")


# ---------------------------------------------------------------- menus

def test_menu_decorators_register_commands():
    from xloil.menus import menu_items

    @xlo.context_menu("Double it", menu=["Cell", "List Range Popup"], disable_screen_updating=True)
    def double_it():
        return 2

    @xlo.menu("Report", menu="Tools", sub_menu="Reports")
    def report():
        """Runs the report"""

    items = {i['command']: i for i in menu_items()}
    assert items["xloMenu_double_it"]['bars'] == ["Cell", "List Range Popup"]
    assert items["xloMenu_report"]['menu'] == "Tools"
    assert items["xloMenu_report"]['tooltip'] == "Runs the report"
    spec = xt.spec(double_it)
    assert "command" in spec.features and spec.name == "xloMenu_double_it"


def test_menu_rejects_unknown_options():
    with pytest.raises(TypeError):
        xlo.menu("X", bogus=True)


# ---------------------------------------------------------------- toolkits

def test_toolkit_registry():
    from xloil.gui.toolkits import (GuiToolkit, register_toolkit, find_toolkit,
                                    registered_toolkits, HwndWindow)

    class Widget: pass

    class MyToolkit(GuiToolkit):
        name = "mine"
        def handles(self, obj):
            return isinstance(obj, Widget)

    register_toolkit(MyToolkit())
    try:
        assert find_toolkit(Widget()).name == "mine"
        assert find_toolkit(HwndWindow(1234)).name == "hwnd"
        assert find_toolkit(object()) is None
        names = [t.name for t in registered_toolkits()]
        assert names[0] == "mine" and {"qt", "tk", "wx", "hwnd"} <= set(names)
    finally:
        from xloil.gui import toolkits
        toolkits._TOOLKITS[:] = [t for t in toolkits._TOOLKITS if t.name != "mine"]


# ---------------------------------------------------------------- plotting

def test_plot_types_registered():
    pytest.importorskip("matplotlib")
    import xloil.plotting as xp
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots()
    assert xp._find_plot_type(ax).name == "matplotlib"
    assert xp._find_plot_type(fig).name == "matplotlib"
    plt.close(fig)


def test_register_custom_plot_type():
    import xloil.plotting as xp
    class MyChart: pass
    xp.register_plot_type(MyChart, to_html=lambda c: "<html></html>", name="mine")
    assert xp._find_plot_type(MyChart()).name == "mine"
    with pytest.raises(ValueError):
        xp.register_plot_type(MyChart)


def test_plot_mode_validation():
    import xloil.plotting as xp
    with pytest.raises(ValueError):
        xp.set_plot_mode("fancy")
    xp.set_plot_mode("static")
    xp.set_plot_mode("auto")


# --------------------------------------------------------- startup / reload

def test_startup_script_runs_once(tmp_path, monkeypatch):
    from xloil import importer
    marker = tmp_path / "ran.txt"
    script = tmp_path / "startup.py"
    script.write_text(f"open(r'{marker}', 'a').write('x')")

    class Addin:
        pathname = str(tmp_path / "a.xll")
        settings_file = str(tmp_path / "a.ini")
        settings = {'xlOil_Python': {'StartupScript': "startup.py"}}

    importer.run_startup_scripts(Addin())
    importer.run_startup_scripts(Addin())
    assert marker.read_text() == "x"


def test_startup_script_list_and_missing(tmp_path):
    from xloil import importer
    class Addin:
        pathname = str(tmp_path / "b.xll")
        settings_file = str(tmp_path / "b.ini")
        settings = {'xlOil_Python': {'StartupScript': ["missing.py"]}}
    importer.run_startup_scripts(Addin())  # logs an error, does not raise


# ----------------------------------------------------------------- CLI

@pytest.fixture
def ini_copy(tmp_path):
    src = Path(__file__).resolve().parents[4] / "config" / "xloil.ini"
    dst = tmp_path / "xlOil.ini"
    dst.write_text(src.read_text(encoding="utf-8"), encoding="utf-8")
    return dst


def test_cli_config_set_get(ini_copy, capsys):
    from xloil import command_line as cl
    cl.config_set("xlOil_Python.LoadModules", '["xloil.xloil_ribbon", "mymod"]', ini=str(ini_copy))
    cl.config_set("xlOil_Python.StartupScript", '"start.py"', ini=str(ini_copy))
    capsys.readouterr()
    cl.config_get("xlOil_Python.LoadModules", ini=str(ini_copy))
    assert "mymod" in capsys.readouterr().out
    import tomlkit
    doc = tomlkit.parse(ini_copy.read_text(encoding="utf-8"))
    assert doc["xlOil_Python"]["StartupScript"] == "start.py"
    # Comments are preserved
    assert "######## xlOil Settings #########" in ini_copy.read_text(encoding="utf-8")


def test_cli_parse_value():
    from xloil.command_line import _parse_value
    assert _parse_value("3") == 3
    assert _parse_value("true") is True
    assert _parse_value("plain text") == "plain text"
    assert list(_parse_value('["a"]')) == ["a"]


def test_cli_help_runs(monkeypatch, capsys):
    from xloil import command_line as cl
    monkeypatch.setattr(sys, "argv", ["xloil"])
    cl.main()
    assert "install" in capsys.readouterr().out


# ------------------------------------------------------------ conda support

def test_python_home_for_conda_venv_and_plain(tmp_path):
    from xloil.command_line import python_home_for
    conda = tmp_path / "miniconda"
    (conda / "conda-meta").mkdir(parents=True)
    (conda / "python.exe").write_text("")
    assert python_home_for(conda / "python.exe") == str(conda)

    env = conda / "envs" / "py311"
    (env / "conda-meta").mkdir(parents=True)
    assert python_home_for(env / "python.exe") == str(env)

    venv = tmp_path / "venv"
    (venv / "Scripts").mkdir(parents=True)
    (venv / "pyvenv.cfg").write_text(f"home = {conda}\ninclude-system-site-packages = false\n")
    assert python_home_for(venv / "Scripts" / "python.exe") == str(conda)

    plain = tmp_path / "Python311"
    plain.mkdir()
    assert python_home_for(plain / "python.exe") is None


def test_ini_python_home_written_and_removed(ini_copy):
    from xloil.command_line import _write_python_path_to_ini, _set_ini_python_home
    import tomlkit
    text = ini_copy.read_text(encoding="utf-8")
    out, ok = _write_python_path_to_ini(text, r"C:\xloil\bin", True,
        replace_paths={'PYTHONEXECUTABLE': r'C:\conda\python.exe', 'PYTHONHOME': r'C:\conda'})
    assert ok
    doc = tomlkit.parse(out)
    env = [e for e in doc['xlOil_Python']['Environment'] if 'PYTHONEXECUTABLE' in e][0]
    assert env['PYTHONHOME'] == r'C:\conda'
    # Idempotent: a second write replaces rather than duplicates
    out2 = _set_ini_python_home(out, r'D:\other')
    assert out2.count("PYTHONHOME='''") == 1 and r"D:\other" in out2
    # Switching to a non-conda python comments it out
    out3 = _set_ini_python_home(out2, None)
    env3 = [e for e in tomlkit.parse(out3)['xlOil_Python']['Environment'] if 'PYTHONEXECUTABLE' in e][0]
    assert 'PYTHONHOME' not in env3


def test_doctor_runs_without_excel(capsys):
    from xloil import command_line as cl
    problems = cl.doctor()
    out = capsys.readouterr().out
    assert "xlOil doctor" in out
    assert isinstance(problems, int)


# ---------------------------------------------------------------- features

def test_feature_toggle_orders_and_keeps_shared_modules():
    from xloil.features import find_feature, set_feature, feature_enabled
    mods = ["xloil.xloil_ribbon", "mymodule"]
    mods = set_feature(mods, find_feature("charts"), True)
    mods = set_feature(mods, find_feature("qt"), True)
    # Qt must load before plotting; user modules stay last
    assert mods.index("xloil.gui.qtpy") < mods.index("xloil.plotting")
    assert mods[0] == "xloil.xloil_ribbon" and mods[-1] == "mymodule"
    # Turning charts off keeps plotting, which qt still needs
    mods = set_feature(mods, find_feature("charts"), False)
    assert "xloil.plotting" in mods and "xloil.matplotlib" not in mods
    assert feature_enabled(find_feature("qt"), mods)
    mods = set_feature(mods, find_feature("qt"), False)
    assert "xloil.plotting" not in mods and "xloil.gui.qtpy" not in mods


def test_feature_catalogue_modules_exist():
    import importlib.util
    from xloil.features import FEATURES
    for f in FEATURES:
        for m in f.modules:
            assert importlib.util.find_spec(m) is not None, m


def test_order_modules_dedupes():
    from xloil.features import order_modules
    assert order_modules(["a", "xloil.plotting", "xloil.gui.qtpy", "a"]) == \
        ["xloil.gui.qtpy", "xloil.plotting", "a"]


def test_setup_subcommand_registered(monkeypatch, capsys):
    from xloil import command_line as cl
    monkeypatch.setattr(sys, "argv", ["xloil", "setup", "--help"])
    with pytest.raises(SystemExit):
        cl.main()
    out = capsys.readouterr().out
    assert "--features" in out and "--yes" in out
