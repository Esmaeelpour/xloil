"""
    Importing this module turns on saving cached objects with workbooks
    (`xloil.cache_persist.enable` with default options). Add
    ``xloil.features.cache_persist`` to LoadModules (or tick it in the ribbon).
"""
from xloil.cache_persist import enable as _enable

_enable()
