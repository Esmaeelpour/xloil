"""
    Runtime behaviour for "stub mode": when xloil is imported outside Excel and
    without the compiled xloil_core extension (e.g. on Linux CI), the generated
    stub module supplies the API surface but every function does nothing.

    This module makes the stubs useful for unit tests by replacing a few parts
    with simple pure-Python implementations:

      * `_FuncSpec` / `_FuncArg` record what the decorators declared
      * `cache` is an in-memory ObjectCache
      * `Caller` returns a configurable fake caller (see `xloil.testing`)
      * `excel_callback` runs the callback immediately
      * `in_wizard` returns False, `_register_functions` records registrations
"""

import concurrent.futures
import itertools
import threading


class FuncArg:
    def __init__(self, name, help="", converter=None, flags=""):
        self.name = name
        self.help = help
        self.converter = converter
        self.flags = flags
        self.default = None

    def __str__(self):
        return f"{self.name}[{self.flags}]"


class FuncSpec:
    def __init__(self, func, args, name="", features=None, help="", category="",
                 local=True, volatile=False, errors=0):
        self.func = func
        self.args = list(args)
        self.name = name
        self.features = features or ""
        self.help = help
        self.category = category
        self.is_local = local
        self.volatile = volatile
        self.errors = errors
        self.return_converter = None

    @property
    def is_async(self):
        return "async" in self.features

    @property
    def is_rtd(self):
        return "rtd" in self.features

    @property
    def is_threaded(self):
        return "threaded" in self.features

    @property
    def is_command(self):
        return "command" in self.features

    @property
    def error_propagation(self):
        return self.errors == 1

    def __str__(self):
        return f"{self.name}({', '.join(a.name for a in self.args)}) [{self.features}]"

    __repr__ = __str__


class ObjectCache:
    """ An in-memory stand-in for xlOil's object cache """
    UNIQUIFIER = "〈"  # same character as the real cache

    def __init__(self):
        self._store = {}
        self._counter = itertools.count()
        self._lock = threading.Lock()

    def add(self, obj, tag="", key=""):
        with self._lock:
            if not key:
                key = f"[stub]{tag or type(obj).__name__}{next(self._counter)}"
            prefix = f"{self.UNIQUIFIER}[{key}"
            n = sum(1 for r in self._store if r.startswith(prefix + ","))
            ref = f"{prefix},{chr(65 + n)}"
            self._store[ref] = obj
            return ref

    __call__ = add

    def get(self, ref, default=None):
        return self._store.get(ref, default)

    def __getitem__(self, ref):
        return self._store[ref]

    def contains(self, ref):
        return ref in self._store

    __contains__ = contains

    def keys(self):
        return list(self._store)

    def remove(self, ref):
        return self._store.pop(ref, None) is not None

    def clear(self):
        self._store.clear()


class FakeCaller:
    """ Mimics xloil.Caller. Configure with `xloil.testing.caller` """
    current = {'workbook': "Book1.xlsx", 'sheet': "Sheet1", 'address': "A1"}

    def __init__(self, *args, **kwargs):
        self._info = dict(FakeCaller.current)

    @property
    def workbook(self):
        return self._info['workbook']

    @property
    def sheet_name(self):
        return self._info['sheet']

    @property
    def range(self):
        return None

    def address(self, style='a1', local=False):
        if local:
            return self._info['address']
        return f"[{self.workbook}]{self.sheet_name}!{self._info['address']}"

    def __str__(self):
        return self.address()


REGISTERED = []          # list of (FuncSpec list, module, addin)
CALLBACKS_RUN = []       # callables passed to excel_callback


def _register_functions(funcs, module=None, addin=None, append=False):
    REGISTERED.append((list(funcs), module, addin))


def _excel_callback(func, wait=0, retry=500, api=''):
    future = concurrent.futures.Future()
    CALLBACKS_RUN.append(func)
    try:
        future.set_result(func())
    except BaseException as e:
        future.set_exception(e)
    return future


class Event:
    """ A minimal event: supports +=, -=, handlers and `fire(*args)` for tests """
    def __init__(self, name):
        self.name = name
        self._handlers = []

    def __iadd__(self, handler):
        self._handlers.append(handler)
        return self

    def __isub__(self, handler):
        self._handlers.remove(handler)
        return self

    add = __iadd__
    remove = __isub__

    def clear(self):
        self._handlers.clear()

    @property
    def handlers(self):
        return tuple(self._handlers)

    def fire(self, *args):
        for handler in list(self._handlers):
            handler(*args)

    def __repr__(self):
        return f"Event({self.name}, {len(self._handlers)} handlers)"


def _install_events(event_module):
    for name in getattr(event_module, "__all__", []):
        if getattr(event_module, name, None) is None and name[0].isupper():
            setattr(event_module, name, Event(name))
    file_events = {}
    def file_change(path, action='modify', subdirs=True):
        return file_events.setdefault((path, action, subdirs), Event(f"file_change:{path}"))
    event_module.file_change = file_change
    event_module.pause = lambda excel=True: None
    event_module.allow = lambda excel=True: None


class LogWriter:
    """ Forwards xloil.log calls to the standard `logging` module ('xloil' logger) """
    _LEVELS = {'trace': 5, 'debug': 10, 'info': 20, 'warn': 30, 'warning': 30, 
               'error': 40, 'critical': 50}

    def __init__(self):
        import logging
        self._logger = logging.getLogger("xloil")

    def _level(self, level):
        return level if isinstance(level, int) else self._LEVELS.get(str(level).lower(), 20)

    def __call__(self, msg, *args, level='info', **kwargs):
        self._logger.log(self._level(level), str(msg), *args)

    def trace(self, msg, *args): self._logger.log(5, str(msg), *args)
    def debug(self, msg, *args): self._logger.debug(str(msg), *args)
    def info(self, msg, *args): self._logger.info(str(msg), *args)
    def warn(self, msg, *args): self._logger.warning(str(msg), *args)
    def warning(self, msg, *args): self._logger.warning(str(msg), *args)
    def error(self, msg, *args): self._logger.error(str(msg), *args)
    def critical(self, msg, *args): self._logger.critical(str(msg), *args)
    def flush(self): pass

    level = 'info'
    path = ''


def install(core_module):
    """ Patches the stub xloil_core module in place """
    _install_events(core_module.event)
    core_module._LogWriter = LogWriter
    core_module._FuncSpec = FuncSpec
    core_module._FuncArg = FuncArg
    core_module.cache = ObjectCache()
    core_module.Caller = FakeCaller
    core_module._register_functions = _register_functions
    core_module.excel_callback = _excel_callback
    core_module.in_wizard = lambda: False
    core_module.XLOIL_STUB_MODE = True
