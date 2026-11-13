import sys
import os

ADDIN_NAME        = "xlOil.xll"
INIFILE_NAME      = "xlOil.ini"
APP_DATA_DIR      = os.path.join(os.environ.get('APPDATA','.'), "xlOil")
XLOIL_INSTALL_DIR = os.path.join(sys.prefix, "share", "xloil")
XLOIL_BIN_DIR     = os.environ.get("XLOIL_BIN_DIR", XLOIL_INSTALL_DIR)

class _SetPathContext:

    def __init__(self, path):
        self._path = path
        self._old_PATH = os.environ['PATH']

    def __enter__(self):
        os.environ['PATH'] += os.pathsep + self._path 
        return self

    def close(self):
        os.environ['PATH'] = self._old_PATH

    def __exit__(self, *args):
        self.close()

def add_dll_path(path):
    """
        Returns a context manager which adds a PATH to the dll search directory
        (either via AddDllDirectory or by changing PATH). The change is reversed
        on context exit
    """
    try:
        return os.add_dll_directory(path)
    except (AttributeError, FileNotFoundError):
        # Either Py < 3.8 or path does not exist (but we have to return a context mgr)
        return _SetPathContext(path)

def _get_environment_strings():
    """
        Does the same as:
            import win32profile
            win32profile.GetEnvironmentStrings()
        But avoids the depedency on pywin32
    """
   
    import ctypes
    import locale

    kernel_func = ctypes.windll.kernel32.GetEnvironmentStringsA
    char_ptr = ctypes.POINTER(ctypes.c_char)
    kernel_func.restype = char_ptr

    # P will point to an block of char formatted as:
    #     name1=val1/0
    #     ...
    #     nameN=valN/0/0
    # (line-breaks added for clarity, they aren't present in the string)
    p = kernel_func()
    try:
    
        result = {}
        start = end = 0
        null = b'\x00' # Null-terminator for C-strings

        while True:
            while p[end] != null:
                end += 1
            if end == start:
                break
                
            keyval = p[start:end].decode(locale.getpreferredencoding()).partition('=')
            # GetEnvironmentStrings returns some strange entries starting with '='
            if any(keyval[0]):
                result[keyval[0]] = keyval[2]
                
            end += 1    # Step over null terminator
            start = end # Move string start pointer
    finally:
        ctypes.windll.kernel32.FreeEnvironmentStringsA(p)
    
    return result

def _pip_qt_installed() -> bool:
    """ True if a pip-wheel Qt binding (which bundles its own Qt DLLs) is installed """
    import importlib.util
    for package, core_dll in (("PySide6", "Qt6Core.dll"), ("PyQt6", os.path.join("Qt6", "bin", "Qt6Core.dll")),
                              ("PySide2", "Qt5Core.dll"), ("PyQt5", os.path.join("Qt5", "bin", "Qt5Core.dll"))):
        try:
            spec = importlib.util.find_spec(package)
        except (ImportError, ValueError):
            spec = None
        if spec is not None and spec.submodule_search_locations:
            for location in spec.submodule_search_locations:
                if os.path.exists(os.path.join(location, core_dll)):
                    return True
    return False


def preload_system_icu():
    r"""
        Avoids a DLL clash between pip-installed Qt and conda environments.

        Pip wheels of Qt link against Windows' own ICU (System32\icuuc.dll). Conda
        environments often contain a different, versioned ICU with the same file
        name in Library\bin, which conda's Python adds to the DLL search path, so
        Qt fails with "DLL load failed ... The specified procedure could not be
        found". Loading the system ICU by full path first makes Windows reuse it
        for Qt. Only acts when both conditions for the clash are present.
    """
    if os.name != "nt":
        return False
    conda_icu = os.path.join(sys.prefix, "Library", "bin", "icuuc.dll")
    if not os.path.exists(conda_icu) or not _pip_qt_installed():
        return False
    import ctypes
    system32 = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32")
    loaded = False
    for name in ("icuuc.dll", "icuin.dll"):
        path = os.path.join(system32, name)
        if os.path.exists(path):
            try:
                ctypes.WinDLL(path)
                loaded = True
            except OSError:
                pass
    return loaded
