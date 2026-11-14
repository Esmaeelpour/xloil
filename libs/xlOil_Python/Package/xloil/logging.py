import traceback
import warnings
from ._core import _LogWriter

log = _LogWriter()

# Python's logging names: log.warning(...) and log.critical(...) work too. Code written with logging habits called
# log.warning, which the core writer does not have: an AttributeError inside Excel only (often in an error path)
def _add_logging_names(cls):
    """Give a log writer class warning() and critical() (as warn() and error()) when it lacks them."""
    for alias, target in (('warning', 'warn'), ('critical', 'error')):
        if not hasattr(cls, alias):
            try:
                setattr(cls, alias, (lambda t: lambda self, msg, *args: getattr(self, t)(msg, *args))(target))
            except (AttributeError, TypeError):
                pass
    return cls


_add_logging_names(_LogWriter)
"""
    An instance of `xloil._LogWriter` which writes a log message to xlOil's log.  The level 
    parameter can be a integer constant from the ``logging`` module or one of the strings
    *error*, *warn*, *info*, *debug* or *trace*.

    Only messages with a level higher than the xlOil log level which is (initially) set
    in the xlOil settings file will be output to the log file. Trace output can only
    be seen with a debug build of xlOil.
"""

def _log_warnings(message, category, filename, lineno, file=None, line=None):
    log(warnings.formatwarning(message, category, filename, lineno, line), level='warn')

from ._core import XLOIL_STUB_MODE
if not XLOIL_STUB_MODE:
    # In stub mode (tests) leave Python's warning display alone
    warnings.showwarning = _log_warnings

def log_except(msg, level='error'):
    """
       Logs '{msg}: {stack trace}' with a default level of 'error'
    """
    log.error(f"{msg}: {traceback.format_exc()}")