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


def test_install_dir_prefers_per_user_binaries(tmp_path, monkeypatch):
    import importlib, site, xloil._paths as paths
    prefix = tmp_path / "conda"; userbase = tmp_path / "user"
    for base in (prefix, userbase):
        (base / "share" / "xloil").mkdir(parents=True)
        (base / "share" / "xloil" / "xlOil.xll").write_text("")
    usersite = userbase / "site-packages"
    monkeypatch.setattr(sys, "prefix", str(prefix))
    monkeypatch.setattr(site, "getuserbase", lambda: str(userbase))
    monkeypatch.setattr(site, "getusersitepackages", lambda: str(usersite))
    # Package in the environment: the environment's binaries win
    monkeypatch.setattr(paths, "__file__", str(prefix / "Lib" / "site-packages" / "xloil" / "_paths.py"))
    assert paths._find_install_dir() == str(prefix / "share" / "xloil")
    # Package installed per-user: the per-user binaries win
    monkeypatch.setattr(paths, "__file__", str(usersite / "xloil" / "_paths.py"))
    assert paths._find_install_dir() == str(userbase / "share" / "xloil")


# ------------------------------------------------------------ config commands

def test_config_env_table_keys(ini_copy, tmp_path, capsys):
    from xloil import command_line as cl
    import tomlkit
    cl.config_set("xlOil_Python.Environment.XLOIL_PYTHON_PATH", "'C:/mine'", ini=str(ini_copy))
    capsys.readouterr()
    cl.config_get("xlOil_Python.Environment.XLOIL_PYTHON_PATH", ini=str(ini_copy))
    assert "C:/mine" in capsys.readouterr().out
    folder = tmp_path / "mods"; folder.mkdir()
    cl.config_add_path(str(folder), ini=str(ini_copy))
    cl.config_add_path(str(folder), ini=str(ini_copy))   # no duplicate
    doc = tomlkit.parse(ini_copy.read_text(encoding="utf-8"))
    env = [t for t in doc["xlOil_Python"]["Environment"] if "XLOIL_PYTHON_PATH" in t][0]
    assert str(env["XLOIL_PYTHON_PATH"]).split(";") == ["C:/mine", str(folder.resolve())]


def test_config_rejects_non_list_modules(ini_copy):
    from xloil import command_line as cl
    with pytest.raises(ValueError):
        cl.config_set("xlOil_Python.LoadModules", "[...]", ini=str(ini_copy))
    with pytest.raises(ValueError):
        cl.config_set("xlOil_Python.LoadModules", "mymodule", ini=str(ini_copy))


def test_config_add_remove_module(ini_copy):
    from xloil import command_line as cl
    import tomlkit
    cl.config_add_module("mytools", ini=str(ini_copy))
    cl.config_add_module("xloil.gui.qtpy", ini=str(ini_copy))
    cl.config_add_module("mytools", ini=str(ini_copy))
    mods = list(tomlkit.parse(ini_copy.read_text(encoding="utf-8"))["xlOil_Python"]["LoadModules"])
    assert mods == ["xloil.xloil_ribbon", "xloil.gui.qtpy", "mytools"]
    cl.config_remove_module("xloil.xloil_ribbon", ini=str(ini_copy))
    mods = list(tomlkit.parse(ini_copy.read_text(encoding="utf-8"))["xlOil_Python"]["LoadModules"])
    assert mods == ["xloil.gui.qtpy", "mytools"]


# --- xlam ribbon --------------------------------------------------------------

@pytest.fixture
def xlam_ribbon(monkeypatch):
    """ xlam_ribbon with ribbon_core replaced: the real one reads Excel's settings on import """
    import types
    import xloil

    class FakeCore(types.ModuleType):
        def __getattr__(self, name):
            fn = lambda ctrl, *args: (name, ctrl.tag, *args)
            setattr(self, name, fn)
            return fn

    fake = FakeCore("xloil.ribbon_core")
    monkeypatch.setitem(sys.modules, "xloil.ribbon_core", fake)
    monkeypatch.setattr(xloil, "ribbon_core", fake, raising=False)
    monkeypatch.delitem(sys.modules, "xloil.xlam_ribbon", raising=False)
    import xloil.xlam_ribbon as mod
    yield mod
    sys.modules.pop("xloil.xlam_ribbon", None)


def test_xlam_bridge_dispatch(xlam_ribbon):
    assert xlam_ribbon.xloRibbonCall("text", "loadmodules") == ("get_load_modules", "loadmodules")
    assert xlam_ribbon.xloRibbonCall("select", "loglevel", "id3", 3.0) == ("set_log_level", "loglevel", "id3", 3)
    assert xlam_ribbon.xloRibbonCall("toggle", "feature:pandas", True) == ("set_feature", "pandas", True)
    # Unknown callbacks are refused (logged), not executed
    assert xlam_ribbon.xloRibbonCall("action", "nope") is None


def test_xlam_bridge_runs_coroutines(xlam_ribbon):
    import asyncio

    async def coro(ctrl, value):
        await asyncio.sleep(0)
        return value * 2
    xlam_ribbon._CALLBACKS[("change", "_test")] = coro
    try:
        assert xlam_ribbon.call("change", "_test", 21) == 42
    finally:
        del xlam_ribbon._CALLBACKS[("change", "_test")]


def test_xlam_ribbon_xml_tags_have_callbacks(xlam_ribbon):
    import re
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).resolve().parents[4] / "tools"))
    import build_xlam
    xml = build_xlam.ribbon_xml()
    from xml.dom.minidom import parseString
    parseString(xml)
    tags = set(re.findall(r'tag="([^"]+)"', xml))
    known = {t for _, t in xlam_ribbon._CALLBACKS}
    missing = [t for t in tags if not t.startswith("feature:") and t not in known]
    assert not missing


def test_ribbon_command_switches_modules(tmp_path, monkeypatch):
    # Runs outside Excel, so it must not import the Excel-only ribbon modules
    for mod in ("xloil.ribbon_core", "xloil.xlam_ribbon", "xloil.xloil_ribbon"):
        monkeypatch.setitem(sys.modules, mod, None)
    from xloil import command_line as cl
    ini = tmp_path / "xlOil.ini"
    ini.write_text('[xlOil_Python]\nLoadModules = ["xloil.xloil_ribbon", "mymod"]\n', encoding="utf-8")
    xlstart = tmp_path / "XLSTART"
    monkeypatch.setattr(cl, "_XL_START_PATH", xlstart)
    monkeypatch.setattr(cl, "_unblock", lambda p: None)
    fake = tmp_path / "xlOil.xlam"
    fake.write_bytes(b"PK")
    monkeypatch.setattr(cl, "_xlam_source", lambda: fake)
    cl.ribbon("xlam", ini=str(ini))
    assert (xlstart / "xlOil.xlam").exists()
    text = ini.read_text()
    assert "xloil.xlam_ribbon" in text and "xloil.xloil_ribbon" not in text and "mymod" in text
    cl.ribbon("off", ini=str(ini))
    assert not (xlstart / "xlOil.xlam").exists()
    assert "ribbon" not in ini.read_text()


# --- vbaProject.bin writer (tools/vba_project.py) -----------------------------

def _tools():
    from pathlib import Path
    tools = str(Path(__file__).resolve().parents[4] / "tools")
    if tools not in sys.path:
        sys.path.insert(0, tools)


def test_vba_compression_round_trip():
    _tools()
    import vba_project as v
    for data in [b"", b"a", b"abc" * 3000, bytes(range(256)) * 20,
                 "Sub X()\r\n    MsgBox \"hi\"\r\nEnd Sub\r\n".encode() * 200]:
        assert v.decompress(v.compress(data)) == data


def test_vba_project_encryption_round_trip():
    _tools()
    import vba_project as v
    for seed in (0x07, 0x0B, 0x0D, 0x00):
        assert v._decrypt(v._encrypt("{A1B2}", b"\x00\xff\x10", seed)) == b"\x00\xff\x10"


def test_vba_project_is_readable_compound_file():
    olefile = pytest.importorskip("olefile")
    import io
    _tools()
    import vba_project as v
    data = v.build_vba_project({"M1": "Sub A()\r\nEnd Sub\r\n"})
    ole = olefile.OleFileIO(io.BytesIO(data))
    streams = {"/".join(s) for s in ole.listdir()}
    assert {"PROJECT", "PROJECTwm", "VBA/dir", "VBA/_VBA_PROJECT", "VBA/M1",
            "VBA/ThisWorkbook", "VBA/Sheet1"} <= streams
    src = v.decompress(ole.openstream("VBA/M1").read())
    assert src.startswith(b'Attribute VB_Name = "M1"') and b"Sub A()" in src


def test_packaged_xlam_found_by_cli():
    from xloil import command_line as cl
    assert cl._xlam_source().exists()


def test_committed_xlam_is_up_to_date():
    import zipfile
    from pathlib import Path
    _tools()
    import build_xlam
    committed = Path(__file__).resolve().parents[1] / "xloil" / "data" / "xlOil.xlam"
    with zipfile.ZipFile(committed) as z:
        for name, data in build_xlam._parts().items():
            data = data.encode("utf-8") if isinstance(data, str) else data
            assert z.read(name) == data, f"{name} differs: run python tools/build_xlam.py"


def test_addins_must_be_signed_reads_policy(monkeypatch):
    """ Simulates the Office policy registry value (QueryValueEx returns (value, type)) """
    from xloil import command_line as cl

    class FakeReg:
        HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE = "HKCU", "HKLM"

        def __init__(self, values):
            self.values = values

        def OpenKey(self, root, path):
            if (root, path) not in self.values:
                raise OSError("no key")
            return (root, path)

        def QueryValueEx(self, key, name):
            return self.values[key], 4

    key = ("HKCU", r"Software\Policies\Microsoft\Office\16.0\Excel\Security")
    monkeypatch.setattr(cl, "reg", FakeReg({key: 1}))
    assert cl.addins_must_be_signed(16) is True
    monkeypatch.setattr(cl, "reg", FakeReg({key: 0}))
    assert cl.addins_must_be_signed(16) is False
    monkeypatch.setattr(cl, "reg", FakeReg({}))
    assert cl.addins_must_be_signed(16) is False


def test_setup_keeps_user_modules(ini_copy, monkeypatch, capsys):
    """ Re-running setup rebuilds LoadModules from the features but keeps your own modules """
    from xloil import command_line as cl
    cl.config_set("xlOil_Python.LoadModules", '["xloil.xloil_ribbon", "xloil.pandas", "mytools"]',
                  ini=str(ini_copy))
    assert cl._user_modules(str(ini_copy)) == ["mytools"]

    captured = {}
    monkeypatch.setattr(cl, "_user_modules", lambda ini=None, _f=cl._user_modules: _f(str(ini_copy)))
    monkeypatch.setattr(cl, "_install_cmd", lambda load_modules=None, **kw: captured.setdefault("mods", load_modules))
    monkeypatch.setattr(cl, "addins_must_be_signed", lambda version=None: True)
    monkeypatch.setattr(cl, "ribbon", lambda kind, ini=None: captured.setdefault("ribbon", kind))
    monkeypatch.setattr(cl, "_check_bitness", lambda: True)
    monkeypatch.setattr(cl, "_excel_installed", lambda: True)
    monkeypatch.setattr(cl, "doctor", lambda: 0)
    cl.setup(features="charts", yes=True, no_pip=True, intellisense=False)
    mods = captured["mods"].split(",")
    assert "mytools" in mods and mods[-1] == "mytools"
    assert "xloil.xlam_ribbon" in mods and "xloil.xloil_ribbon" not in mods
    assert captured["ribbon"] == "xlam"


def test_load_modules_ignores_comments(ini_copy):
    """ The default ini's comments mention xloil.xloil_ribbon: doctor must only look at
        the LoadModules value, or it reports the COM ribbon after `xloil ribbon xlam` """
    from xloil import command_line as cl
    assert "xloil.xloil_ribbon" in ini_copy.read_text(encoding="utf-8").split("LoadModules=")[0]
    cl.config_set("xlOil_Python.LoadModules", '["xloil.xlam_ribbon", "mytools"]', ini=str(ini_copy))
    assert cl._load_modules(str(ini_copy)) == ["xloil.xlam_ribbon", "mytools"]


# ---- lessons from a workbook module in real Excel (VDOT bridge workbook, 2026-11-13) ----------------------------------

def test_log_has_logging_names():
    """log.warning / log.critical work, as in Python's logging (the core writer had only warn: AttributeError in Excel)"""
    from xloil.logging import _add_logging_names
    xlo.log.warning("a warning %s", 1)
    xlo.log.critical("a critical %s", 2)

    class Core:                                     # like the compiled writer: warn and error only
        def __init__(self):
            self.got = []
        def warn(self, msg, *args):
            self.got.append(("warn", msg % args))
        def error(self, msg, *args):
            self.got.append(("error", msg % args))
    _add_logging_names(Core)
    w = Core()
    w.warning("w %s", 1)
    w.critical("c %s", 2)
    assert w.got == [("warn", "w 1"), ("error", "c 2")]


def test_decorated_function_keeps_its_module():
    """Copying __module__ from a decorated function gives the function's own module (was xloil.register, never scanned)"""
    @xlo.func
    def keeps_module(x: float) -> float:
        return x
    assert keeps_module.__module__ == __name__
    assert keeps_module.__qualname__.endswith("keeps_module")


def test_linked_workbook_from_a_helper(tmp_path):
    """A helper in another module, called while a workbook module runs, finds the workbook up the call stack; a module
    can also be named"""
    import types
    from xloil import importer
    wb_mod = tmp_path / "Book1.py"
    wb_mod.write_text("import helper_lw\nFOUND = helper_lw.which()\n")
    (tmp_path / "helper_lw.py").write_text("import xloil\ndef which():\n    return xloil.linked_workbook()\n")
    importer._linked_workbooks[str(wb_mod)] = "Book1.xlsx"
    sys.path.insert(0, str(tmp_path))
    try:
        ns = {"__file__": str(wb_mod), "__name__": "wbmod_test"}
        exec(compile(wb_mod.read_text(), str(wb_mod), "exec"), ns)
        assert ns["FOUND"] == "Book1.xlsx"
        assert xlo.linked_workbook(types.SimpleNamespace(__file__=str(wb_mod))) == "Book1.xlsx"
        assert xlo.linked_workbook(str(wb_mod)) == "Book1.xlsx"
        assert xlo.linked_workbook() is None          # this test module is not a workbook module
    finally:
        sys.path.remove(str(tmp_path))
        importer._linked_workbooks.pop(str(wb_mod), None)


def test_name_errors_fixed_after_workbook_module_loads(monkeypatch):
    """After a workbook module registers its functions, its workbook's #NAME? cells are re-entered (a full recalculation
    does not clear them)"""
    from xloil import importer, com, _core
    queued, fixed = [], []
    monkeypatch.setattr(xloil_core, "excel_callback", lambda fn, wait=0, **k: queued.append((fn, wait)), raising=False)
    monkeypatch.setattr(com, "fix_name_errors", lambda wb: fixed.append(wb))
    monkeypatch.setattr(_core, "workbooks", {"Bridge.xlsx": "the workbook"})
    importer._fix_name_errors_later(r"C:\work\Bridge.xlsx")
    assert len(queued) == 1 and queued[0][1] > 0
    queued[0][0]()
    assert fixed == ["the workbook"]


def test_excel_array_like_excel():
    """Rows as Excel hands them to an Array argument: code testing 'rows or []' fails on it, as in Excel"""
    a = xt.excel_array([["key", 1.5, None], ["k2", "x", 3]])
    assert a.shape == (2, 3) and a.dtype == object
    with pytest.raises(ValueError):
        bool(a or [])


def test_remove_takes_the_ribbon_too(tmp_path, monkeypatch, capsys):
    """xloil remove / uninstall takes the add-in and the xlOil.xlam ribbon out of XLSTART (the xlam calls xlOil: left
    alone it errored at every Excel start); IntelliSense stays (named); a file Excel holds gives a clear message"""
    import xloil.command_line as cl
    from xloil.intellisense import default_xll_name
    xlstart = tmp_path / "XLSTART"
    xlstart.mkdir()
    for n in ("xlOil.xll", "xlOil.xlam", default_xll_name(), "other.xlam"):
        (xlstart / n).write_bytes(b"x")
    monkeypatch.setattr(cl, "_XL_START_PATH", xlstart)
    monkeypatch.setattr(cl, "_XLL_INSTALL_PATH", xlstart / "xlOil.xll")
    monkeypatch.setattr(cl, "_find_excel_version", lambda: 16)
    monkeypatch.setattr(cl, "_remove_addin", lambda version, path: None)
    cl._remove_xloil()
    assert sorted(p.name for p in xlstart.iterdir()) == sorted([default_xll_name(), "other.xlam"])
    assert "intellisense remove" in capsys.readouterr().out
    (xlstart / "xlOil.xll").write_bytes(b"x")
    real_remove = cl.os.remove
    def locked(path):
        raise PermissionError(13, "in use")
    monkeypatch.setattr(cl.os, "remove", locked)
    with pytest.raises(SystemExit, match="close every Excel window|Close every Excel window"):
        cl._remove_xloil()
    monkeypatch.setattr(cl.os, "remove", real_remove)


def test_clean_waits_for_pip_and_removes_logs(tmp_path, monkeypatch):
    """xloil clean deletes the settings and every xlOil log, then waits for pip (its output came after the prompt)"""
    import subprocess
    import xloil.command_line as cl
    app = tmp_path / "xlOil"
    app.mkdir()
    for n in ("xlOil.ini", "xlOil.log", "xlOil.1234.log", "xlOil.ini.before-debug.bak"):
        (app / n).write_text("x")
    monkeypatch.setattr(cl, "APP_DATA_DIR", str(app))
    monkeypatch.setattr(cl, "_remove_xloil", lambda: None)
    ran = []
    monkeypatch.setattr(subprocess, "run", lambda cmd, **k: ran.append(cmd))
    cl._clean_xloil()
    assert sorted(p.name for p in app.iterdir()) == ["xlOil.ini.before-debug.bak"]
    assert ran and ran[0][-3:] == ["uninstall", "--yes", "xloil"]


def test_setup_keeps_the_chosen_ribbon(ini_copy, monkeypatch):
    """Re-running setup (an upgrade) keeps the ribbon chosen before (xlam or off); a first install gets the COM ribbon;
    the signed-add-ins policy always gets xlam (0.22.1.post13 reset an xlam choice to COM and left xlOil.xlam behind)"""
    from xloil import command_line as cl
    assert cl._ribbon_kind(False, None) == "com"
    assert cl._ribbon_kind(False, "com") == "com"
    assert cl._ribbon_kind(False, "xlam") == "xlam"
    assert cl._ribbon_kind(False, "off") == "off"
    assert cl._ribbon_kind(True, "off") == "xlam"
    for mods, want in (('["xloil.xlam_ribbon", "xloil.pandas"]', "xlam"), ('["xloil.xloil_ribbon"]', "com"),
                       ('["xloil.pandas"]', "off")):
        cl.config_set("xlOil_Python.LoadModules", mods, ini=str(ini_copy))
        assert cl._ribbon_choice(str(ini_copy)) == want
    assert cl._ribbon_choice(str(ini_copy.parent / "missing.ini")) is None
    cl.config_set("xlOil_Python.LoadModules", '["xloil.xlam_ribbon", "mytools"]', ini=str(ini_copy))
    captured = {}
    monkeypatch.setattr(cl, "_ribbon_choice", lambda ini=None, _f=cl._ribbon_choice: _f(str(ini_copy)))
    monkeypatch.setattr(cl, "_user_modules", lambda ini=None, _f=cl._user_modules: _f(str(ini_copy)))
    monkeypatch.setattr(cl, "_install_cmd", lambda load_modules=None, **kw: captured.setdefault("mods", load_modules))
    monkeypatch.setattr(cl, "addins_must_be_signed", lambda version=None: False)
    monkeypatch.setattr(cl, "ribbon", lambda kind, ini=None: captured.setdefault("ribbon", kind))
    monkeypatch.setattr(cl, "_check_bitness", lambda: True)
    monkeypatch.setattr(cl, "_excel_installed", lambda: True)
    monkeypatch.setattr(cl, "doctor", lambda: 0)
    cl.setup(features="pandas", yes=True, no_pip=True, intellisense=False)
    assert captured["ribbon"] == "xlam" and "xloil.xlam_ribbon" in captured["mods"].split(",")
    assert "xloil.xloil_ribbon" not in captured["mods"].split(",") and "mytools" in captured["mods"].split(",")


def test_setup_keeps_the_chosen_features(ini_copy, monkeypatch):
    """Re-running setup with --yes (the installer's -Yes upgrade) keeps the features the settings already turn on;
    a first install gets pandas and charts (0.22.1.post14 replaced a Qt choice with pandas and charts)"""
    from xloil import command_line as cl
    assert cl._feature_defaults(str(ini_copy.parent / "missing.ini")) == {"pandas", "charts"}
    cl.config_set("xlOil_Python.LoadModules", '["xloil.xloil_ribbon", "xloil.gui.qtpy", "xloil.plotting"]',
                  ini=str(ini_copy))
    assert cl._feature_defaults(str(ini_copy)) == {"qt"}
    cl.config_set("xlOil_Python.LoadModules", '["xloil.xloil_ribbon"]', ini=str(ini_copy))
    assert cl._feature_defaults(str(ini_copy)) == {"pandas", "charts"}
    cl.config_set("xlOil_Python.LoadModules", '["xloil.xloil_ribbon", "xloil.gui.qtpy", "xloil.plotting"]',
                  ini=str(ini_copy))
    captured = {}
    for name in ("_ribbon_choice", "_user_modules", "_feature_defaults"):
        monkeypatch.setattr(cl, name, lambda ini=None, _f=getattr(cl, name): _f(str(ini_copy)))
    monkeypatch.setattr(cl, "_install_cmd", lambda load_modules=None, **kw: captured.setdefault("mods", load_modules))
    monkeypatch.setattr(cl, "addins_must_be_signed", lambda version=None: False)
    monkeypatch.setattr(cl, "ribbon", lambda kind, ini=None: None)
    monkeypatch.setattr(cl, "_check_bitness", lambda: True)
    monkeypatch.setattr(cl, "_excel_installed", lambda: True)
    monkeypatch.setattr(cl, "doctor", lambda: 0)
    import xloil.features as feats
    monkeypatch.setattr(feats, "feature_available", lambda f: True)
    cl.setup(yes=True, no_pip=True, intellisense=False)
    mods = captured["mods"].split(",")
    assert "xloil.gui.qtpy" in mods and "xloil.plotting" in mods
    assert "xloil.pandas" not in mods and "xloil.matplotlib" not in mods


def test_config_set_takes_a_list_powershell_stripped(ini_copy):
    """Windows PowerShell 5.1 strips the inner double quotes it passes to a program: xloil config set ...
    '["a", "b"]' arrives as [a, b], which was rejected; bare module names in brackets are taken as that list"""
    from xloil import command_line as cl
    cl.config_set("xlOil_Python.LoadModules", "[xloil.xloil_ribbon, xloil.gui.qtpy, my_tools]", ini=str(ini_copy))
    assert cl._load_modules(str(ini_copy)) == ["xloil.xloil_ribbon", "xloil.gui.qtpy", "my_tools"]
    cl.config_set("xlOil_Python.LoadModules", "[]", ini=str(ini_copy))
    assert cl._load_modules(str(ini_copy)) == []
    with pytest.raises(ValueError, match="square brackets"):
        cl.config_set("xlOil_Python.LoadModules", "xloil.pandas", ini=str(ini_copy))
    with pytest.raises(ValueError, match="square brackets"):
        cl.config_set("xlOil_Python.LoadModules", "[a b, c]", ini=str(ini_copy))


def test_one_line_installer_uses_iex_of_irm():
    """The documented one-liner is `iex (irm URL)`: `irm URL | iex` failed in a conda-initialised Windows PowerShell 5.1
    ("The terminator '#>' is missing from the multiline comment"), 3 runs of 3; `iex (irm URL)` ran 3 of 3.
    -NoProfile: the profile's conda hook re-activated base in the child shell, so an activated conda env was ignored
    and xlOil went into base"""
    root = Path(__file__).resolve().parents[4]
    for rel in ("tools/install.ps1", "docs/source/xlOil_Python/Installing.rst"):
        p = root / rel
        if not p.exists():
            pytest.skip("not in the source tree")
        text = p.read_text(encoding="utf-8")
        assert ('powershell -NoProfile -ExecutionPolicy Bypass -c "iex (irm '
                'https://github.com/Esmaeelpour/xloil/releases/latest/download/install.ps1)"') in text, rel
        assert 'install.ps1 | iex"' not in text, rel


def test_installer_help_block_closes_once():
    """install.ps1's <# ... #> help block must not contain the closing marker itself (a quoted error message once
    ended it early and broke the whole script)"""
    p = Path(__file__).resolve().parents[4] / "tools" / "install.ps1"
    if not p.exists():
        pytest.skip("not in the source tree")
    text = p.read_text(encoding="utf-8")
    assert text.startswith("<#") and text.count("#>") == 1 and text.count("<#") == 1


def test_xloattr_calls_builtin_methods():
    """=xloAttr("abc","upper") returned the method uncalled (built-ins are neither ismethod nor isfunction);
    routines are called with the given arguments, other attributes and classes are returned as they are"""
    from xloil import excelfuncs as ef
    assert ef._xlo_attr_helper("python works", "upper") == "PYTHON WORKS"
    assert ef._xlo_attr_helper("a,b", "split", ",") == ["a", "b"]
    assert ef._xlo_attr_helper({"k": 1}, "get", "k") == 1
    assert ef._xlo_attr_helper(dt.date(2026, 1, 2), "isoformat") == "2026-01-02"
    assert ef._xlo_attr_helper(dt.date(2026, 1, 2), "year") == 2026
    import collections
    assert ef._xlo_attr_helper(collections, "OrderedDict") is collections.OrderedDict


def test_setup_without_excel_says_so_before_installing(monkeypatch, capsys):
    """On a computer without Excel, setup pip-installed pandas and matplotlib, then install_xloil crashed with a
    FileNotFoundError traceback (found by the install smoke test on a clean Windows runner). Now it stops first, with
    a message; `xloil install` on its own gives the same message instead of a traceback"""
    import subprocess
    import xloil.command_line as cl
    monkeypatch.setattr(cl, "_check_bitness", lambda: True)
    monkeypatch.setattr(cl, "excel_path", lambda: None)
    def no_excel():
        raise FileNotFoundError(2, "The system cannot find the file specified")
    monkeypatch.setattr(cl, "_find_excel_version", no_excel)
    ran = []
    monkeypatch.setattr(subprocess, "run", lambda cmd, **k: ran.append(cmd))
    assert cl._excel_installed() is False
    assert cl.setup(features="pandas", yes=True, intellisense=False) == 1
    assert "Excel was not found" in capsys.readouterr().out
    assert ran == []                                    # nothing pip-installed
    with pytest.raises(SystemExit, match="Excel was not found"):
        cl.install_xloil()
    monkeypatch.setattr(cl, "_find_excel_version", lambda: 16)
    assert cl._excel_installed() is True
