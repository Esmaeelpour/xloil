"""
    Saves cached Python objects inside the workbook and restores them on open.

    Normally xlOil's object cache lives only in memory: when a workbook is closed
    and re-opened, cells holding cache references (e.g. ``<[Book1]Sheet1!R1C1,A``)
    point at nothing until the functions which created them recalculate. For
    expensive objects - a fitted model, a large dataset - that recalculation can
    be slow or impossible (the source data may have gone).

    With persistence enabled, picklable objects are serialised into the workbook
    when it is saved and put back into the cache, under the *same reference
    strings*, when it is opened. Downstream formulas then work immediately.

    Enable it once, e.g. in a startup module or your add-in module::

        import xloil.cache_persist
        xloil.cache_persist.enable()

    then choose which objects to save: either per function::

        @xlo.func(persist=True)
        def fit_model(data: xlo.Array(float)):
            return MyModel().fit(data)   # cached and saved with the workbook

    or by type::

        xloil.cache_persist.persist_type(MyModel)

    or everything picklable: ``enable(persist_all=True)``.

    Objects are stored in a Custom XML Part (base64-encoded pickles). Only load
    workbooks from trusted sources with this feature enabled: unpickling can run
    arbitrary code. Set ``trusted_only=True`` (the default) to only restore objects
    into workbooks whose location Excel trusts (no Protected View).
"""

import base64
import pickle
import threading
from typing import Callable, Optional, Dict

import xloil_core
from .logging import log, log_except
from .metadata import encode_store, decode_store, _read_xml_part, _write_xml_part, _workbook_com

CACHE_NAMESPACE = "urn:xloil:cache:1"

_lock = threading.Lock()
_marked_ids = set()
_persist_types = set()
_settings = {
    'persist_all': False,
    'dumps': lambda obj: pickle.dumps(obj, protocol=pickle.HIGHEST_PROTOCOL),
    'loads': pickle.loads,
    'enabled': False,
    'max_bytes': 200 * 1024 * 1024,
}


def mark_persistent(obj):
    """
        Marks an object so that, if it is in the object cache when its workbook
        is saved, it is saved too. Used by ``@xlo.func(persist=True)``; can be
        called directly on objects you add with `xloil.cache.add`.
    """
    if obj is None or isinstance(obj, (int, float, str, bool)):
        return obj
    with _lock:
        _marked_ids.add(id(obj))
    return obj


def persist_type(*types):
    """ Persists all cached objects which are instances of any of the given types """
    _persist_types.update(types)


def _should_persist(obj) -> bool:
    if _settings['persist_all']:
        return True
    if id(obj) in _marked_ids:
        return True
    return bool(_persist_types) and isinstance(obj, tuple(_persist_types))


def _ref_workbook(ref: str) -> Optional[str]:
    """ Cache refs look like U[BookName]Sheet!R1C1,A: extract BookName """
    if len(ref) < 4 or ref[1] != '[':
        return None
    end = ref.find(']', 2)
    return ref[2:end] if end > 0 else None


def _ref_to_key(ref: str) -> str:
    """ The cache's `add(key=...)` prepends 'U[' and appends ',X' """
    return ref[2:-2]


def _iter_cache_refs(cache):
    """
        Yields (reference string, object) for every object in the cache. In 
        xlOil 0.22, `ObjectCache.keys()` returns keys without the ",A" counter
        suffix used in the reference strings written to cells, so probe the
        suffixes when a key is not itself a valid reference.
    """
    for key in cache.keys():
        if len(key) > 2 and key[-2] == ',' and key in cache:
            yield key, cache.get(key)
            continue
        for i in range(64):
            ref = f"{key},{chr(65 + i)}"
            obj = cache.get(ref)
            if obj is None:
                break
            yield ref, obj


def collect(workbook_name: str) -> Dict[str, bytes]:
    """
        Serialises the persistable cache objects belonging to the named workbook.
        Returns {ref: pickled_bytes}. Objects which fail to pickle are skipped
        with a warning.
    """
    cache = xloil_core.cache
    result = {}
    total = 0
    for ref, obj in _iter_cache_refs(cache):
        if _ref_workbook(ref) != workbook_name:
            continue
        if not _should_persist(obj):
            continue
        try:
            data = _settings['dumps'](obj)
        except Exception as e:
            log.warn(f"cache_persist: cannot save {ref} ({type(obj).__name__}): {e}")
            continue
        total += len(data)
        if total > _settings['max_bytes']:
            log.warn(f"cache_persist: size limit {_settings['max_bytes']} bytes exceeded, "
                     f"remaining objects in {workbook_name} not saved")
            break
        result[ref] = data
    return result


def save(workbook=None):
    """ Writes the persistable cache objects into the workbook (does not save the file) """
    wb = _workbook_com(workbook)
    entries = collect(wb.Name)
    _write_xml_part(wb, CACHE_NAMESPACE,
                    encode_store(entries, CACHE_NAMESPACE, "xloilCache") if entries else None)
    log.info(f"cache_persist: stored {len(entries)} objects in {wb.Name}")
    return len(entries)


def restore(workbook=None, overwrite: bool = False) -> int:
    """
        Loads objects saved in the workbook back into the cache under their
        original reference strings. Returns the number restored.
    """
    wb = _workbook_com(workbook)
    if _settings.get('trusted_only', True):
        try:
            # Workbooks in Protected View do not appear in Application.Workbooks,
            # but check explicitly in case of future changes
            if wb.Application.ProtectedViewWindows.Count > 0:
                for i in range(1, wb.Application.ProtectedViewWindows.Count + 1):
                    if wb.Application.ProtectedViewWindows(i).Workbook.FullName == wb.FullName:
                        return 0
        except Exception:
            pass

    entries = decode_store(_read_xml_part(wb, CACHE_NAMESPACE), CACHE_NAMESPACE)
    cache = xloil_core.cache
    count = 0
    # Sort so ,A comes before ,B for refs from the same cell
    for ref in sorted(entries):
        if not overwrite and ref in cache:
            continue
        try:
            obj = _settings['loads'](entries[ref])
        except Exception as e:
            log.warn(f"cache_persist: cannot restore {ref}: {e}")
            continue
        new_ref = cache.add(obj, key=_ref_to_key(ref))
        mark_persistent(obj)
        if new_ref != ref:
            log.warn(f"cache_persist: {ref} restored as {new_ref}")
        count += 1
    log.info(f"cache_persist: restored {count} objects into {wb.Name}")
    return count


def _on_before_save(wb_name, save_as_ui, cancel):
    try:
        save(wb_name)
    except Exception:
        log_except(f"cache_persist: failed saving cache objects for {wb_name}")


def _on_open(wb_path, wb_name):
    def work():
        try:
            restore(wb_name)
        except Exception:
            log_except(f"cache_persist: failed restoring cache objects for {wb_name}")
    # Run after the open event completes, so the workbook is fully available
    xloil_core.excel_callback(work)


def enable(persist_all: bool = False,
           dumps: Callable = None,
           loads: Callable = None,
           trusted_only: bool = True,
           max_bytes: int = 200 * 1024 * 1024,
           restore_open_workbooks: bool = True):
    """
        Turns on saving of cached objects with workbooks.

        Parameters
        ----------
        persist_all: bool
            If True, every picklable cached object is saved. Otherwise only objects
            from ``@func(persist=True)`` functions, marked with `mark_persistent`,
            or of a type given to `persist_type`.
        dumps, loads: callable
            Serialiser functions, default `pickle.dumps` / `pickle.loads`. For
            example pass ``cloudpickle.dumps`` or a ``dill`` equivalent to handle
            lambdas and locally defined classes.
        trusted_only: bool
            Do not restore objects into workbooks open in Protected View
        max_bytes: int
            Upper limit on the serialised size stored per workbook
        restore_open_workbooks: bool
            Also restore objects into workbooks which are already open
    """
    _settings['persist_all'] = persist_all
    _settings['trusted_only'] = trusted_only
    _settings['max_bytes'] = max_bytes
    if dumps is not None:
        _settings['dumps'] = dumps
    if loads is not None:
        _settings['loads'] = loads

    if _settings['enabled']:
        return
    _settings['enabled'] = True

    xloil_core.event.WorkbookBeforeSave += _on_before_save
    xloil_core.event.WorkbookOpen += _on_open

    if restore_open_workbooks:
        def restore_all():
            for wb in xloil_core.app().workbooks:
                try:
                    restore(wb.name)
                except Exception:
                    log_except(f"cache_persist: restore failed for {wb.name}")
        xloil_core.excel_callback(restore_all)


def disable():
    """ Stops saving / restoring cached objects """
    if not _settings['enabled']:
        return
    xloil_core.event.WorkbookBeforeSave -= _on_before_save
    xloil_core.event.WorkbookOpen -= _on_open
    _settings['enabled'] = False
