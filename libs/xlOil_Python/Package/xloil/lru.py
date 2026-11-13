"""
    A least-recently-used cache decorator for worksheet functions.

    `functools.lru_cache` cannot be used directly with most Excel functions
    because Excel passes arrays (numpy arrays, lists, dicts, ExcelArray views),
    which are unhashable. `xloil.lru_cache` converts arguments to a hashable key
    first, so::

        @xlo.func
        @xlo.lru_cache(maxsize=256)
        def slow_calc(prices: xlo.Array(float), window: int = 20):
            ...

    returns the stored result if the same values are passed again. Put
    `xloil.func` *outside* (above) `lru_cache` so xlOil sees the original
    signature and annotations.

    Cached results are shared, so do not mutate a returned object.
"""

import functools
import threading
import time
from collections import OrderedDict
from typing import Any, Callable, NamedTuple, Optional

import numpy as np


class CacheInfo(NamedTuple):
    hits: int
    misses: int
    maxsize: Optional[int]
    currsize: int


class _Unhashable(Exception):
    pass


def make_key(value: Any) -> Any:
    """
        Converts a value into a hashable equivalent. Arrays are keyed by dtype,
        shape and contents; lists/tuples recursively; dicts by sorted items.
        Raises `_Unhashable` if no key can be made.
    """
    if value is None or isinstance(value, (bool, int, float, str, bytes)):
        # Keep int 1 and float 1.0 distinct: Excel functions can behave differently
        return (type(value).__name__, value)

    if isinstance(value, np.ndarray):
        if value.dtype == object:
            return ('ndarray_o', value.shape, tuple(make_key(x) for x in value.ravel()))
        return ('ndarray', value.dtype.str, value.shape, value.tobytes())

    if isinstance(value, np.generic):
        return ('npscalar', value.dtype.str, value.item())

    if isinstance(value, (list, tuple)):
        return (type(value).__name__, tuple(make_key(x) for x in value))

    if isinstance(value, dict):
        return ('dict', tuple(sorted(((make_key(k), make_key(v)) for k, v in value.items()),
                                     key=repr)))

    # xloil.ExcelArray is a view of Excel's memory which may be gone later
    to_numpy = getattr(value, "to_numpy", None)
    if callable(to_numpy) and type(value).__name__ == "ExcelArray":
        return make_key(to_numpy())

    try:
        import pandas as pd
        if isinstance(value, (pd.DataFrame, pd.Series)):
            hashed = pd.util.hash_pandas_object(value, index=True).values
            columns = tuple(map(str, value.columns)) if isinstance(value, pd.DataFrame) else value.name
            return ('pandas', type(value).__name__, value.shape, columns, hashed.tobytes())
    except ImportError:
        pass

    try:
        hash(value)
        return ('obj', type(value).__qualname__, value)
    except TypeError:
        pass

    try:
        import pickle
        return ('pickle', pickle.dumps(value, protocol=pickle.HIGHEST_PROTOCOL))
    except Exception:
        raise _Unhashable(type(value).__name__)


def lru_cache(maxsize: Optional[int] = 128, typed: bool = True, ttl: Optional[float] = None):
    """
        Decorator: caches results of the wrapped function keyed on its argument
        values, discarding the least recently used entries beyond `maxsize`.

        Parameters
        ----------
        maxsize: int or None
            Maximum entries; None means unbounded
        typed: bool
            Accepted for compatibility with `functools.lru_cache`. Keys are always
            typed (1 and 1.0 are different) because Excel conversions may differ
        ttl: float
            If given, entries older than this many seconds are recomputed

        The wrapper has `cache_info()` and `cache_clear()` methods like
        `functools.lru_cache`. If arguments cannot be turned into a key the
        function is simply called without caching.
    """
    if callable(maxsize):
        # Used as @lru_cache without parentheses
        return lru_cache()(maxsize)

    def decorate(fn: Callable):
        lock = threading.RLock()
        store: "OrderedDict[Any, tuple]" = OrderedDict()
        stats = {'hits': 0, 'misses': 0}

        @functools.wraps(fn)
        def wrapped(*args, **kwargs):
            try:
                key = (make_key(args), make_key(tuple(sorted(kwargs.items()))))
            except _Unhashable:
                return fn(*args, **kwargs)

            now = time.monotonic()
            with lock:
                entry = store.get(key)
                if entry is not None and (ttl is None or now - entry[1] < ttl):
                    store.move_to_end(key)
                    stats['hits'] += 1
                    return entry[0]
                stats['misses'] += 1

            # Compute outside the lock so slow functions don't serialise each other
            result = fn(*args, **kwargs)

            with lock:
                store[key] = (result, now)
                store.move_to_end(key)
                if maxsize is not None:
                    while len(store) > maxsize:
                        store.popitem(last=False)
            return result

        def cache_info() -> CacheInfo:
            with lock:
                return CacheInfo(stats['hits'], stats['misses'], maxsize, len(store))

        def cache_clear():
            with lock:
                store.clear()
                stats['hits'] = stats['misses'] = 0

        wrapped.cache_info = cache_info
        wrapped.cache_clear = cache_clear
        return wrapped

    return decorate
