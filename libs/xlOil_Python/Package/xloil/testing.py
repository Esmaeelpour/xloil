"""
    Helpers for unit testing xlOil code without Excel.

    Outside Excel on a non-Windows platform (or with the environment variable
    ``XLOIL_STUBS=1``), ``import xloil`` loads *stub mode*: the full API can be
    imported, decorators work and record what they declare, and nothing talks to
    Excel. This lets you run tests for your add-in in any CI system::

        # test_myaddin.py  (run with: XLOIL_STUBS=1 pytest)
        import xloil.testing as xt
        import myaddin

        def test_add():
            assert xt.call(myaddin.add, 1, 2) == 3

        def test_declared():
            spec = xt.spec(myaddin.add)
            assert spec.name == "add" and [a.name for a in spec.args] == ["x", "y"]

        def test_ticker():
            assert xt.call(myaddin.ticker, limit=3) == [0, 1, 2]   # RTD generator

        def test_caller_dependent():
            with xt.caller(workbook="Prices.xlsx", sheet="Data", address="B4"):
                assert xt.call(myaddin.where_am_i) == "[Prices.xlsx]Data!B4"

    The ``xloil-stubs`` package on PyPI-style indexes (built from this repo)
    installs just the Python part of xlOil for this purpose; do not install it
    in the same environment as the full ``xloil`` package.
"""

import asyncio
import contextlib
import inspect
from typing import Any, List

from ._core import XLOIL_STUB_MODE, XLOIL_EMBEDDED

import xloil_core


def in_stub_mode() -> bool:
    """ True if xloil was imported in stub mode (no Excel / compiled core) """
    return bool(getattr(xloil_core, "XLOIL_STUB_MODE", False))


def _require_stub_mode(what):
    if not in_stub_mode():
        raise RuntimeError(f"{what} is only available in stub mode (set XLOIL_STUBS=1)")


def unwrap(fn):
    """ Returns the plain Python function behind an `xloil.func` declaration """
    return getattr(fn, "__wrapped__", fn)


def spec(fn):
    """ Returns the function spec recorded by `xloil.func` (name, args, help, features) """
    found = getattr(fn, "_xloil_spec", None)
    if found is None:
        raise ValueError(f"{fn!r} is not decorated with xloil.func")
    return found


def call(fn, *args, **kwargs) -> Any:
    """
        Calls an xlOil-declared function as Python, bypassing type conversion.
        Coroutines are run to completion; generators and async generators (RTD
        functions) are exhausted and the list of yielded values returned. Pass
        a ``limit`` keyword... no: to stop an infinite generator, use `take`.
    """
    target = unwrap(fn)
    result = target(*args, **kwargs)
    if inspect.iscoroutine(result):
        return asyncio.run(result)
    if inspect.isasyncgen(result):
        async def collect():
            return [x async for x in result]
        return asyncio.run(collect())
    if inspect.isgenerator(result):
        return list(result)
    return result


def take(fn, n: int, *args, **kwargs) -> List[Any]:
    """ Returns the first `n` values yielded by a (possibly infinite) RTD generator function """
    target = unwrap(fn)
    result = target(*args, **kwargs)
    if inspect.isasyncgen(result):
        async def collect():
            out = []
            async for x in result:
                out.append(x)
                if len(out) >= n:
                    break
            await result.aclose()
            return out
        return asyncio.run(collect())
    out = []
    for x in result:
        out.append(x)
        if len(out) >= n:
            break
    result.close()
    return out


@contextlib.contextmanager
def caller(workbook: str = "Book1.xlsx", sheet: str = "Sheet1", address: str = "A1"):
    """ Within this context, `xloil.Caller()` reports the given location """
    _require_stub_mode("caller()")
    from .stubs._runtime import FakeCaller
    previous = dict(FakeCaller.current)
    FakeCaller.current = {'workbook': workbook, 'sheet': sheet, 'address': address}
    try:
        yield
    finally:
        FakeCaller.current = previous


def registered(module=None):
    """
        Returns the function specs registered (via `xloil.register_functions` or
        module scanning) in stub mode, optionally filtered by module.
    """
    _require_stub_mode("registered()")
    from .stubs._runtime import REGISTERED
    out = []
    for specs, mod, _ in REGISTERED:
        if module is None or mod is module:
            out.extend(specs)
    return out


def pending(module) -> list:
    """ Returns the specs `xloil.func` has queued for registration in `module` """
    from .register import _LANDMARK_TAG
    return list(getattr(module, _LANDMARK_TAG, ()))


def fire_event(name: str, *args):
    """ Fires an xlOil event (e.g. 'WorkbookOpen') in stub mode, calling its handlers """
    _require_stub_mode("fire_event()")
    getattr(xloil_core.event, name).fire(*args)


def reset():
    """ Clears the stub cache, registrations and recorded callbacks """
    _require_stub_mode("reset()")
    from .stubs import _runtime
    xloil_core.cache.clear()
    _runtime.REGISTERED.clear()
    _runtime.CALLBACKS_RUN.clear()
