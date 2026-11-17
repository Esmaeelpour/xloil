"""SheetControl with a toolkit that embeds its own widget (the Qt path) and one that does not (the Win32 path), with
stand-ins for Excel and the toolkit: run anywhere with XLOIL_STUBS=1 (no Excel, no Qt).

    XLOIL_STUBS=1 PYTHONPATH=libs/xlOil_Python/Package python libs/xlOil_Python/tests_py/test_sheet_control_embed.py
"""
import concurrent.futures as futures
import sys
import types

import xloil  # noqa: F401
import xloil.gui.sheet_control as S

FAIL = []


def check(what, ok, info=''):
    print(('OK    ' if ok else 'FAIL  ') + what + ('' if ok else f'  ({info})'))
    if not ok:
        FAIL.append(what)


class _Now(futures.Executor):                      # runs at once (the toolkit thread, here the same thread)
    def submit(self, fn, *a, **k):
        f = futures.Future()
        f.set_result(fn(*a, **k))
        return f


class Kit(S.GuiToolkit):
    name = 'fake'

    def __init__(self, embeds):
        self.embeds, self.calls = embeds, []

    def handles(self, obj):
        return True

    def executor(self):
        return _Now()

    def create(self, obj):
        return obj

    def get_hwnd(self, widget):
        self.calls.append('get_hwnd')
        return 777, True

    def embed(self, widget, parent_hwnd):
        self.calls.append(('embed', parent_hwnd))
        return 555 if self.embeds else None

    def place(self, widget, x, y, w, h):
        self.calls.append(('place', x, y, w, h))

    def set_visible(self, widget, visible):
        self.calls.append(('visible', visible))

    def detach(self, widget):
        self.calls.append('detach')

    def destroy(self, widget):
        self.calls.append('destroy')


class Shape:
    Name, Left, Top, Width, Height = 'xloPlot_t', 10.0, 20.0, 300.0, 200.0


class Shapes:
    def __init__(self):
        self.s = None

    def Item(self, n):
        if self.s is None:
            raise KeyError(n)
        return self.s

    def AddShape(self, *a):
        self.s = Shape()
        return self.s


class Pane:
    def PointsToScreenPixelsX(self, p):
        return 100 + p * 1.5                       # 150 % display scaling

    def PointsToScreenPixelsY(self, p):
        return 50 + p * 1.5


class Panes:
    Count = 1

    def Item(self, k):
        return Pane()


class Window:
    def __init__(self, sheet):
        self.sheet, self.Visible, self.WindowState, self.Hwnd, self.Panes = sheet, True, 0, 1, Panes()
        self.Parent = sheet.Parent

    @property
    def ActiveSheet(self):
        return self.sheet


class Book:
    Name = 'B.xlsx'

    def Windows(self, k):
        return self.win


class Sheet:
    def __init__(self):
        self.Name, self.Parent, self.Shapes = '19 Cap', Book(), Shapes()
        self.Parent.Application = types.SimpleNamespace(ActiveWindow=None)
        self.Parent.win = Window(self)

    def Range(self, a):
        return types.SimpleNamespace(Left=10.0, Top=20.0, Width=60.0, Height=15.0)


S.find_grid_window = lambda h: 4242
S._user32 = types.SimpleNamespace(IsWindow=lambda h: True, ShowWindowAsync=lambda *a: None, SetWindowPos=lambda *a: None,
                                  ScreenToClient=lambda g, p: None, SetParent=lambda *a: None)
S._GetWindowLongPtr = lambda *a: 0
S._SetWindowLongPtr = lambda *a: 0
S._Tracker.instance = classmethod(lambda cls: types.SimpleNamespace(ensure_running=lambda: None))
S.wt = types.SimpleNamespace(POINT=lambda x, y: types.SimpleNamespace(x=x, y=y))
S.ctypes = types.SimpleNamespace(byref=lambda p: p)

kit = Kit(embeds=True)
sh = Sheet()
c = S.SheetControl('w', sheet=sh, at='B2', size=(300, 200), name='xloPlot_t', toolkit=kit)
check('Qt-like toolkit: embedded through the toolkit in the grid window, never re-parented by Win32',
      ('embed', 4242) in kit.calls and 'get_hwnd' not in kit.calls and c.hwnd == 555, str(kit.calls))
places = [x for x in kit.calls if isinstance(x, tuple) and x[0] == 'place']
check('placed over its placeholder in physical pixels (150 % scaling: 10 pt -> x 115, 300 pt wide -> 450)',
      places and places[-1] == ('place', 115, 80, 450, 300), str(places))
n = len(kit.calls)
c.update_position()
check('not moved again while the placeholder has not moved', len(kit.calls) == n, str(kit.calls[n:]))
sh.Parent.win.sheet = types.SimpleNamespace(Name='Other')
c.update_position()
check('hidden when another sheet is shown', kit.calls[-1] == ('visible', False), str(kit.calls[-1:]))
sh.Parent.win.sheet = sh
c.update_position()
check('shown again in place on its sheet', kit.calls[-1][0] == 'place', str(kit.calls[-1:]))
c.delete()
check('deleted: detached from Excel before it is destroyed', kit.calls[-2:] == ['detach', 'destroy'], str(kit.calls[-2:]))

kit2 = Kit(embeds=False)
c2 = S.SheetControl('w', sheet=Sheet(), at='B2', size=(300, 200), name='xloPlot_u', toolkit=kit2)
check('a toolkit that does not embed: the Win32 way as before', 'get_hwnd' in kit2.calls and c2.hwnd == 777, str(kit2.calls))

# ---- clipped to the scrolling pane (frozen rows, sheet tabs, scroll bars) --------------------------------------------
check('clip: wholly inside -> no region', S._clip_region((0, 0, 999, 999), (50, 50, 200, 200)) is None)
check('clip: top cut by frozen rows', S._clip_region((0, 100, 999, 999), (50, 50, 200, 200)) == (0, 50, 200, 200))
check('clip: scrolled out of the pane -> empty region', S._clip_region((0, 0, 10, 10), (50, 50, 200, 200)) == (0, 0, 0, 0))


class VPane(Pane):                                 # the cells the pane shows start at 30 pt (frozen rows above)
    VisibleRange = types.SimpleNamespace(Left=0.0, Top=30.0, Width=600.0, Height=300.0)   # last row partly hidden


win = types.SimpleNamespace(DisplayHorizontalScrollBar=True, DisplayWorkbookTabs=True, DisplayVerticalScrollBar=True)
rect = types.SimpleNamespace(left=0, top=0, right=700, bottom=300)
u32 = S._user32
S._user32 = types.SimpleNamespace(GetClientRect=lambda g, r: r.__dict__.update(rect.__dict__),
                                  ClientToScreen=lambda g, p: p.__dict__.update(x=p.x + 100, y=p.y + 50),
                                  GetDpiForWindow=lambda g: 144, GetSystemMetricsForDpi=lambda i, dpi: 25)
S.wt.RECT = lambda: types.SimpleNamespace(left=0, top=0, right=0, bottom=0)
area = S._pane_area(win, VPane(), 4242)
check('pane area: from the first scrolling row to above the tabs / left of the scroll bar',
      area == (100, 95, 775, 325), str(area))
S._user32 = u32

regions = []
S._set_region = lambda hwnd, clip: regions.append((hwnd, clip))
S._pane_area = lambda window, pane, grid: (100, 150, 2000, 2000)       # frozen rows end at y 150
kit3 = Kit(embeds=True)
c3 = S.SheetControl('w', sheet=Sheet(), at='B2', size=(300, 200), name='xloPlot_v', toolkit=kit3)
check('embedded view clipped to the pane on its toolkit thread (placed at y 80, pane starts at 150)',
      regions and regions[-1] == (555, (0, 70, 450, 300)), str(regions))
n = len(regions)
c3.update_position()
check('region not set again while nothing moved', len(regions) == n, str(regions[n:]))
S._pane_area = lambda window, pane, grid: (0, 0, 5000, 5000)
c3.update_position()
check('region removed when the view is wholly inside the pane', regions[-1] == (555, None), str(regions[-1:]))

# ---- Excel's quit: controls go before the GUI thread, and Excel never waits for ever ---------------------------------
import threading
import time
import xloil.gui as G

done = threading.Thread(target=lambda: None)
done.start()
check('join: a finished thread joins', G._join_pumping(done, 1.0))
stop = threading.Event()
stuck = threading.Thread(target=stop.wait, daemon=True)
stuck.start()
t0 = time.monotonic()
check('join: a thread that never ends is left after the timeout', not G._join_pumping(stuck, 0.3)
      and time.monotonic() - t0 < 2, f'{time.monotonic() - t0:.1f} s')
stop.set()
f = futures.Future()
threading.Timer(0.1, lambda: f.set_result(7)).start()
check('wait_pumping returns the result', G.wait_pumping(f, 2) == 7)

order = []
kit4 = Kit(embeds=True)
kit4.calls = order
S.SheetControl('w', sheet=Sheet(), at='B2', size=(300, 200), name='xloPlot_w', toolkit=kit4)
del order[:]
fake = types.SimpleNamespace(_ready=types.SimpleNamespace(result=lambda: True), _thread=done, _shutdown='quit',
                             submit=lambda fn, *a: order.append(('submit', fn)))
G._GuiExecutor.shutdown(fake)
check('shutdown: sheet controls detached before the GUI thread is told to quit',
      'detach' in order and order.index('detach') < order.index(('submit', 'quit')), str(order))

# ---- the Qt binding with WebEngine is chosen when qtpy's first choice has none ---------------------------------------
import ast
import os
import tempfile

src = open(os.path.join(os.path.dirname(G.__file__), 'qtpy.py'), encoding='utf-8').read()
fn = next(n for n in ast.parse(src).body if isinstance(n, ast.FunctionDef) and n.name == '_choose_binding')
ns = {'os': os, 'sys': types.SimpleNamespace(modules={}), 'log': types.SimpleNamespace(info=lambda m: None)}
exec(compile(ast.Module([fn], []), 'qtpy.py', 'exec'), ns)
root = tempfile.mkdtemp()
for pkg, files in (('PyQt5', ['QtCore.pyd']), ('PySide6', ['QtCore.pyd', 'QtWebEngineWidgets.pyd'])):
    os.makedirs(os.path.join(root, pkg))
    for name in files:
        open(os.path.join(root, pkg, name), 'w').close()
import importlib.util as IU
real_find = IU.find_spec
IU.find_spec = lambda name: (types.SimpleNamespace(submodule_search_locations=[os.path.join(root, name)])
                             if name in ('PyQt5', 'PySide6') else None)
saved = os.environ.pop('QT_API', None)
try:
    ns['_choose_binding']()
    check('binding: PySide6 (has WebEngine) over PyQt5 (has not)', os.environ.get('QT_API') == 'pyside6',
          os.environ.get('QT_API'))
    os.environ['QT_API'] = 'pyqt5'
    ns['_choose_binding']()
    check('binding: QT_API set by the user is kept', os.environ['QT_API'] == 'pyqt5')
finally:
    IU.find_spec = real_find
    os.environ.pop('QT_API', None)
    if saved is not None:
        os.environ['QT_API'] = saved

print('\nALL PASS' if not FAIL else f'\n{len(FAIL)} FAILED')
sys.exit(1 if FAIL else 0)
