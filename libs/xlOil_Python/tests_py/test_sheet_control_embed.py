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
print('\nALL PASS' if not FAIL else f'\n{len(FAIL)} FAILED')
sys.exit(1 if FAIL else 0)
