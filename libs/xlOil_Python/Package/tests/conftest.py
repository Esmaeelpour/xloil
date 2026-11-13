"""
    Pytest configuration for the stub-mode unit tests. These tests exercise the
    pure-Python parts of xlOil without Excel, so they run on any platform:

        cd libs/xlOil_Python/Package
        pytest tests
"""
import os
import sys
from pathlib import Path

os.environ.setdefault("XLOIL_STUBS", "1")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pytest


class FakeCom:
    """
        A permissive stand-in for a COM object: records attribute writes and
        returns child FakeComs for unknown attributes / calls.
    """
    def __init__(self, name="com", **attrs):
        object.__setattr__(self, "_name", name)
        object.__setattr__(self, "_children", {})
        object.__setattr__(self, "_writes", [])
        for k, v in attrs.items():
            object.__setattr__(self, k, v)

    def __getattr__(self, item):
        children = object.__getattribute__(self, "_children")
        if item not in children:
            children[item] = FakeCom(f"{self._name}.{item}")
        return children[item]

    def __setattr__(self, key, value):
        self._writes.append((key, value))
        object.__setattr__(self, key, value)

    def __call__(self, *args, **kwargs):
        key = f"call{args}"
        children = object.__getattribute__(self, "_children")
        if key not in children:
            children[key] = FakeCom(f"{self._name}{args}")
        return children[key]

    def writes(self):
        return dict(self._writes)


class FakeRange(FakeCom):
    """ Tracks Resize/Offset so formatting code can be checked """
    def __init__(self, row=1, col=1, nrows=1, ncols=1, log=None):
        super().__init__(f"R{row}C{col}:{nrows}x{ncols}")
        object.__setattr__(self, "row", row)
        object.__setattr__(self, "col", col)
        object.__setattr__(self, "nrows", nrows)
        object.__setattr__(self, "ncols", ncols)
        object.__setattr__(self, "log", log if log is not None else [])
        self.log.append(self)

    def Resize(self, nrows, ncols):
        return FakeRange(self.row, self.col, nrows, ncols, self.log)

    def Offset(self, rows, cols):
        return FakeRange(self.row + rows, self.col + cols, self.nrows, self.ncols, self.log)

    def Rows(self, i):
        return FakeRange(self.row + i - 1, self.col, 1, self.ncols, self.log)

    @property
    def Columns(self):
        rng = self
        class _Cols:
            def __call__(_, j):
                return FakeRange(rng.row, rng.col + j - 1, rng.nrows, 1, rng.log)
            def AutoFit(_):
                rng.log.append(("autofit", rng.ncols))
        return _Cols()

    def ClearFormats(self):
        self.log.append(("clear", self.row, self.col))


class FakeXmlPart:
    def __init__(self, parts, xml):
        self._parts = parts
        self.XML = xml

    def Delete(self):
        self._parts.remove(self)


class FakeXmlParts:
    """ Mimics Workbook.CustomXMLParts """
    def __init__(self):
        self.parts = []

    def Add(self, xml):
        part = FakeXmlPart(self.parts, xml)
        self.parts.append(part)
        return part

    def SelectByNamespace(self, ns):
        matching = [p for p in self.parts if f'xmlns="{ns}"' in p.XML]
        class _Sel:
            Count = len(matching)
            def Item(_, i):
                return matching[i - 1]
        return _Sel()


class FakeWorkbook:
    def __init__(self, name="Book1.xlsx"):
        self.Name = name
        self.FullName = f"C:\\temp\\{name}"
        self.CustomXMLParts = FakeXmlParts()


@pytest.fixture
def fake_range():
    return FakeRange()


@pytest.fixture
def fake_workbook():
    return FakeWorkbook()


@pytest.fixture(autouse=True)
def reset_stubs():
    import xloil.testing as xt
    xt.reset()
    yield
