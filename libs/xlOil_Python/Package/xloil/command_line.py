import sys
import os
from pathlib import Path
import re
try:
    import winreg as reg
except ImportError:  # Not Windows: config commands still work, install does not
    reg = None
from contextlib import suppress
import binascii
import shutil as sh
from ._paths import *
from argparse import ArgumentParser

_XL_START_PATH = Path(os.getenv('APPDATA', '.')) / "Microsoft" / "Excel" / "XLSTART"
_XLL_INSTALL_PATH = _XL_START_PATH / ADDIN_NAME

def _excel_regpath(version):
    return rf"Software\Microsoft\Office\{version}\Excel"

def _find_excel_version():
    # Will fetch a value like "Excel.Application.16"
    # Note for Office 2019, it still gives version 16...hmm
    verstr = reg.QueryValue(reg.HKEY_CLASSES_ROOT, r"Excel.Application\CurVer")
    return int(verstr.replace('Excel.Application.',''))


def _check_VBA_access(version):

    user_access = 1
    lm_access = 1

    with suppress(OSError): 
        user_key = reg.OpenKey(reg.HKEY_CURRENT_USER, _excel_regpath(version) + r"\Security")
        user_access, _ = reg.QueryValueEx(user_key, "AccessVBOM")
    with suppress(OSError):
        lm_key =  reg.OpenKey(reg.HKEY_LOCAL_MACHINE, _excel_regpath(version) + r"\Security")
        lm_access, _ = reg.QueryValueEx(lm_key, "AccessVBOM")

    if user_access == 0 or lm_access == 0:
        print("To ensure xlOil local functions work, allow access to the VBA Object Model in\n" +
			"Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings\n")


def _get_xloil_bin_dir():

    # We look in the a possibly overriden (by an env var) bin dir, the normal install 
    # path and the current directory
    for bin_path in [Path(XLOIL_BIN_DIR), Path(XLOIL_INSTALL_DIR), Path(".")]:
        if (bin_path / ADDIN_NAME).exists():
            return bin_path

    raise Exception(f"{ADDIN_NAME} not found")

def _remove_from_resiliancy(filename, version):

    # Source https://stackoverflow.com/questions/751048/

    #Converts the File Name string to UTF16 Hex
    filename_hex = binascii.hexlify(filename.encode('utf-16'))

    # If we can't find the key or exit the for loop, suppress the error
    with suppress(OSError): 
        regkey = reg.OpenKey(reg.HKEY_CURRENT_USER, _excel_regpath(version) + r"\Resiliency\DisabledItems")

        for i in range(1024):
            name, value, = reg.EnumValue(regkey, i)
            value_hex = binascii.hexlify(value.encode('utf-16'))
            if filename_hex in value:
                reg.DeleteValue(regkey, name)


def _remove_addin(excel_version, addin_path):

    # If we can't find the key or exit the for loop, suppress the error
    with suppress(OSError): 
        regkey = reg.OpenKey(reg.HKEY_CURRENT_USER, _excel_regpath(excel_version) + r"\Add-in Manager")

        # Cycles through all the properties and delete if it contains the file name.
        for i in range(1024):
            name, value, = reg.EnumValue(regkey, i)
            if addin_path in value:
                reg.DeleteValue(regkey, name)


def _toml_lit_string(s:str):
    # TOML literal strings have a lot of quotes and escapes, this function does the encoding
    return "'''" + s.replace('\\','\\\\') + "'''"

def _venv_home(exe: Path):
    """ Returns the 'home' directory from a venv's pyvenv.cfg, or None if not a venv """
    for cfg in [exe.parent / "pyvenv.cfg", exe.parent.parent / "pyvenv.cfg"]:
        if cfg.exists():
            for line in cfg.read_text(encoding='utf-8', errors='replace').splitlines():
                key, sep, value = line.partition('=')
                if sep and key.strip().lower() == 'home':
                    return Path(value.strip())
    return None


def python_home_for(executable) -> str:
    """
        Returns the PYTHONHOME xlOil must set for the given python.exe, or None 
        if none is needed.

        Conda distributions (base installs and envs) do not let an embedded 
        interpreter locate the standard library from python.exe alone: Python 
        fails at start-up with "failed to get the Python codec of the filesystem
        encoding". Setting PYTHONHOME to the installation directory fixes this.
        For a venv created from a conda Python, the home is the base install
        named in pyvenv.cfg; the venv itself is still picked up via 
        PYTHONEXECUTABLE. python.org installs need nothing.
    """
    exe = Path(executable)
    base = _venv_home(exe) or exe.parent
    if (base / "conda-meta").is_dir():
        return str(base)
    return None


def _get_python_paths():
    """
        Returns the paths to be set in the xlOil.ini file (the appropriate stubs
        must already exist).
    """
    return { 
        'PYTHONEXECUTABLE': sys.executable,
        'PYTHONHOME': python_home_for(sys.executable),
    }


def _set_ini_python_home(ini_txt: str, home) -> str:
    """
        Writes (or comments out, if `home` is None) a PYTHONHOME line in the
        [xlOil_Python] environment block, next to PYTHONEXECUTABLE.
    """
    active = re.compile(r'^(\s*)PYTHONHOME\s*=.*$', re.M)
    if home is None:
        return active.sub(lambda m: m.group(1) + "#" + m.group(0).lstrip(), ini_txt)
    line_value = "PYTHONHOME=" + "'''" + str(home) + "'''"
    if active.search(ini_txt):
        return active.sub(lambda m: m.group(1) + line_value, ini_txt, count=1)
    exe_line = re.compile(r'^(\s*)PYTHONEXECUTABLE\s*=.*$', re.M)
    m = exe_line.search(ini_txt)
    if m is None:
        return ini_txt
    comment = m.group(1) + "# Needed for conda pythons, see xloil.command_line.python_home_for\n"
    return ini_txt[:m.end()] + "\n" + comment + m.group(1) + line_value + ini_txt[m.end():]

def _write_python_path_to_ini(ini_txt, bin_dir:str, comment_reg_keys:bool, replace_paths=None):

    if replace_paths is None:
        replace_paths = _get_python_paths()
    replace_paths = dict(replace_paths)
    if 'PYTHONEXECUTABLE' in replace_paths and 'PYTHONHOME' not in replace_paths:
        replace_paths['PYTHONHOME'] = python_home_for(replace_paths['PYTHONEXECUTABLE'])
    python_home = replace_paths.pop('PYTHONHOME', None)

    fails = 0

    def replace(pat, repl):
        nonlocal ini_txt
        ini_txt = re.sub(pat, repl, ini_txt, count=1, flags=re.M)
            
    def check_replace(pat, repl):
        nonlocal fails, ini_txt
        if re.search(pat, ini_txt, flags=re.M) is None:
            print(f"Failed to match pattern {pat}")
            fails += 1
        else:
            replace(pat, repl)
    
    for var, value in replace_paths.items():
        check_replace(r'^(\s*' + var + r'\s*=).*', r'\g<1>' + _toml_lit_string(value))
        
    # Set XLOIL_PATH
    check_replace(r'^(\s*XLOIL_PATH\s*=).*', r'\g<1>' + _toml_lit_string(str(bin_dir)))

    ini_txt = _set_ini_python_home(ini_txt, python_home)
    if python_home:
        print(f"Conda python detected: setting PYTHONHOME={python_home}")
    
    # Comment out the now usused code to get the python paths from the registry
    # Don't error if this fails as it's not critical
    if comment_reg_keys:
        
        for key in ["xlOil_PythonRegKey"]:
            replace(rf'^(\s*{key}\s*=.*)', r'#\g<1>')

    return ini_txt, fails == 0
    
   
def _unblock(path):
    """
        Removes the Windows "Mark of the Web" (Zone.Identifier stream) from a file.
        Office refuses to load XLL add-ins marked as downloaded from the internet,
        e.g. when the binaries came from a browser-downloaded release zip.
    """
    try:
        os.remove(str(path) + ":Zone.Identifier")
        print(f"Unblocked {path} (removed internet download mark)")
    except OSError:
        pass


def excel_path():
    """ Returns the path of EXCEL.EXE from the registry, or None """
    if reg is None:
        return None
    for root in (reg.HKEY_LOCAL_MACHINE, reg.HKEY_CURRENT_USER):
        with suppress(OSError):
            key = reg.OpenKey(root, r"Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe")
            value, _ = reg.QueryValueEx(key, "")
            if value and Path(value).exists():
                return Path(value)
    return None


def excel_bitness():
    """ Returns 64, 32 or None (unknown) for the installed Excel """
    exe = excel_path()
    if exe is None:
        return None
    try:
        import struct
        with open(exe, 'rb') as f:
            f.seek(0x3C)
            pe_offset = struct.unpack('<I', f.read(4))[0]
            f.seek(pe_offset + 4)
            machine = struct.unpack('<H', f.read(2))[0]
        return {0x8664: 64, 0x14c: 32, 0xAA64: 64}.get(machine)
    except OSError:
        return None


def _check_bitness():
    import struct
    py_bits = struct.calcsize("P") * 8
    xl_bits = excel_bitness()
    if xl_bits is not None and xl_bits != py_bits:
        print(f"WARNING: Excel is {xl_bits}-bit but this Python is {py_bits}-bit. The add-in "
              f"will not load: install a {xl_bits}-bit Python and the matching xlOil build.")
        return False
    return True


def _configure_bundled_ini(bin_dir: Path):
    """
        Writes this Python's paths into the xlOil.ini shipped next to the binaries
        if it has none. The core add-in reads that file when no %APPDATA% settings
        exist, e.g. when only `xloil create` has been run on this computer.
    """
    ini = bin_dir / INIFILE_NAME
    try:
        text = ini.read_text(encoding='utf-8')
    except OSError:
        return
    if re.search(r"^\s*PYTHONEXECUTABLE\s*=\s*'''\s*\S", text, flags=re.M) and \
            not re.search(r"^\s*PYTHONEXECUTABLE\s*=\s*'''<", text, flags=re.M):
        return
    text, ok = _write_python_path_to_ini(text, bin_dir, comment_reg_keys=True)
    if ok:
        with suppress(OSError):
            ini.write_text(text, encoding='utf-8')
            print(f"Configured {ini} for {sys.executable}")


def install_xloil(ini_template:str=None, 
                  addin_name=None,
                  replace_ini=False, 
                  replace_paths=None):

    ini_path = Path(APP_DATA_DIR) / INIFILE_NAME

    if addin_name is None:
        addin_name = ADDIN_NAME
        
    excel_version = _find_excel_version()

    # Just in case we got put in Excel's naughty corner for misbehaving addins
    _remove_from_resiliancy(ADDIN_NAME, excel_version)

    # Check access to the VBA Object model (for local functions)
    _check_VBA_access(excel_version)

    # Ensure XLSTART dir really exists (its parents too: on a new user profile
    # where Excel has never run, %APPDATA%\Microsoft\Excel doesn't exist yet)
    _XL_START_PATH.mkdir(parents=True, exist_ok=True)

    bin_dir = _get_xloil_bin_dir()

    _check_bitness()

    # Copy the XLL
    try:
        sh.copy(bin_dir / ADDIN_NAME, _XLL_INSTALL_PATH)
    except PermissionError:
        raise SystemExit(f"Can't replace {_XLL_INSTALL_PATH}: Excel is probably running. "
                         "Close every Excel window (check Task Manager for EXCEL.EXE) and try again.")
    _unblock(_XLL_INSTALL_PATH)
    for f in bin_dir.iterdir():
        if f.suffix.lower() in ('.dll', '.pyd', '.xll'):
            _unblock(f)
    print("Installed ", _XLL_INSTALL_PATH)
    
    # Copy the ini file to APPDATA, avoiding overwriting any existing ini
    if not replace_ini and ini_path.exists():
        print("Found existing settings file at \n", ini_path)
    else:
        ini_path.parent.mkdir(parents=True, exist_ok=True)
        sh.copy(ini_template or bin_dir / INIFILE_NAME, ini_path)

    # Edit the xloil.ini file. To preserve comments and whitespace it's easier to just use
    # regex replace rather than read the file as structured TOML
    ini_txt = ini_path.read_text(encoding='utf-8')

    ini_txt, success = _write_python_path_to_ini(ini_txt, bin_dir, 
                                                 comment_reg_keys=True, 
                                                 replace_paths=replace_paths)

    # Check if any of the counts is not 1, i.e. the expression matched zero or multiple times
    if not success:
        print(f'WARNING: Failed to set python paths in {ini_path}. You may have to do this manually.')
    else:
        ini_path.write_text(ini_txt, encoding='utf-8')
        print(f'Edited {ini_path} to point to {sys.prefix} python distribution.')

def _remove_xloil():
    """
        Removes xlOil from Excel: the add-in and the xlOil.xlam ribbon from XLSTART (the ribbon's macros call
        xlOil, so left alone it showed errors at every Excel start), and the add-in's registry entries. The settings
        file stays. The IntelliSense add-in, which other add-ins can use too, stays: `xloil intellisense remove`.
    """
    try:
        excel_version = _find_excel_version()
    except OSError:
        excel_version = None
    if excel_version is not None:
        try:
            _remove_addin(excel_version, _XLL_INSTALL_PATH)   # Ensure no xlOil addins are in the registry
        except OSError:
            pass
    for path in (_XLL_INSTALL_PATH, _XL_START_PATH / XLAM_NAME):
        try:
            os.remove(path)
            print("Removed ", path)
        except FileNotFoundError:
            pass
        except PermissionError:
            raise SystemExit(f"Can't remove {path}: Excel is probably running. Close every Excel window "
                             "(check Task Manager for EXCEL.EXE) and try again.")
    from .intellisense import default_xll_name
    if (_XL_START_PATH / default_xll_name()).exists():
        print(f"The IntelliSense add-in stays in {_XL_START_PATH}: `xloil intellisense remove` removes it")


def _clean_xloil():
    """ Removes xlOil from Excel, deletes its settings file and logs, then uninstalls the package (pip). """
    _remove_xloil()
    app_dir = Path(APP_DATA_DIR)
    for path in [app_dir / INIFILE_NAME] + sorted(app_dir.glob("xlOil*.log")):
        try:
            path.unlink()
            print("Removed ", path)
        except FileNotFoundError:
            pass
        except PermissionError:
            print(f"Can't remove {path}: close Excel and delete it by hand")
    import subprocess
    import sys

    # Wait for pip, so its output comes before the prompt returns
    subprocess.run([sys.executable, "-m", "pip", "uninstall", "--yes", "xloil"])


def _create_addin(filename):

    basename = Path(os.path.splitext(filename)[0])

    xll_path = basename.with_suffix(".xll")
    ini_path = basename.with_suffix(".ini")

    bin_dir = _get_xloil_bin_dir()
    print("xlOil binaries found at:", str(bin_dir))

    sh.copy(bin_dir / ADDIN_NAME,    xll_path)
    sh.copy(bin_dir / INIFILE_NAME,  ini_path)
    _unblock(xll_path)
    _configure_bundled_ini(bin_dir)
    
    print("New addin created at:", xll_path)

    # Edit ini file
    ini_txt = ini_path.read_text(encoding='utf-8')
    
    # Assume we want the xlOil_Python plugin as we're running a python script
    ini_txt, count = re.subn(r'^(\s*Plugins\s*=).*', r'\g<1>["xlOil_Python"]', ini_txt, flags=re.M)

    # Assume we want the python paths set to the distribution running this script
    ini_txt, success = _write_python_path_to_ini(ini_txt, bin_dir, 
                                                 comment_reg_keys=True)
    
    ini_path.write_text(ini_txt, encoding='utf-8')

    print("xlOil_Python plugin enabled using python installed at: ", sys.prefix)
 
    
def _package_pyinstaller(ini_template: str, 
                         makespec: bool = False, 
                         extra_args: str = None):

    if ini_template is None:
        ini_template = Path(APP_DATA_DIR) / INIFILE_NAME
    else:
        ini_template = Path(ini_template)
        
    python3_dll = os.path.join(sys.base_exec_prefix, "python3.dll")

    # Use the filename stems as when the script is run, these files will be
    # available in its current directory. 
    ini_filename = ini_template.name
    addin_filename = "xloil.xll"
    
    entry_point = "install_main.py"
    
    install_stub = \
        'import os\n' \
        'import sys\n' \
        'from xloil.command_line import install_xloil\n' \
        f'install_xloil(ini_template="{ini_filename}", addin_name="{addin_filename}",\n' \
        '   replace_paths={"PYTHONEXECUTABLE": os.path.join(sys.prefix, "python.exe")})\n'
    
    Path(entry_point).write_text(install_stub)

    args = [
        '',
        entry_point,
        '--onedir',
        '--debug=noarchive',
        '--nowindow',
        f'--add-data={ini_template}:.',
        f'--add-data={python3_dll}:.',
        f'--add-data={sys.executable}:.',
    ]

    if extra_args is not None:
        args += extra_args
       
    print(args)
    
    try:
        if makespec:
            from PyInstaller.utils.cliutils.makespec import run
        else:
            from PyInstaller.__main__ import run
    except ImportError:
        raise ImportError("To run `xloil package` you need the pyinstaller package installed")

    sys.argv = args
    run()


#
# Configuration commands: status / config / intellisense
#

def _ini_path(ini=None) -> Path:
    return Path(ini) if ini else Path(APP_DATA_DIR) / INIFILE_NAME


def _load_ini(ini=None):
    import tomlkit
    path = _ini_path(ini)
    if not path.exists():
        raise FileNotFoundError(f"Settings file {path} not found: run `xloil install` first")
    return path, tomlkit.parse(path.read_text(encoding='utf-8'))


def _split_key(key: str):
    parts = key.split('.')
    if len(parts) < 2:
        raise ValueError("Keys look like Section.Name, e.g. xlOil_Python.LoadModules")
    return parts


def _parse_value(text: str):
    """ Parses a value written on the command line as TOML, falling back to a string """
    import tomlkit
    try:
        item = tomlkit.parse(f"v = {text}\n")["v"]
        # Return a plain python value so tomlkit formats it afresh (an item taken
        # from another document keeps that document's whitespace)
        return item.unwrap() if hasattr(item, "unwrap") else item
    except Exception:
        return text


# Settings which must be lists of strings: catching a typo here stops xlOil
# failing to load at the next Excel start
_LIST_SETTINGS = {"LoadModules", "Plugins", "DateFormats", "AutoReloadExcludePaths"}


def _is_table_array(node) -> bool:
    return isinstance(node, list) and len(node) > 0 and all(hasattr(x, "items") for x in node)


def _resolve(doc, parts, create=False):
    """
        Walks `parts` (all but the final key) through the settings document and
        returns the table that holds the final key. Arrays of tables, such as the
        [[xlOil_Python.Environment]] blocks, are searched for the table that
        already contains the final key (or indexed with a number, e.g.
        Environment.0). If no table has the key, the last one is used.
    """
    import tomlkit
    node = doc
    path, final = parts[:-1], parts[-1]
    for i, part in enumerate(path):
        if _is_table_array(node):
            node = node[int(part)]
            continue
        if part not in node:
            if not create:
                raise KeyError(".".join(parts[:i + 1]))
            node[part] = tomlkit.table()
        node = node[part]
    if _is_table_array(node):
        node = next((t for t in node if final in t), node[-1])
    return node, final


def config_get(key: str, ini: str = None):
    """
        Prints the value of a setting, e.g. ``xloil config get xlOil_Python.LoadModules``
        or ``xloil config get xlOil_Python.Environment.XLOIL_PYTHON_PATH``
    """
    _, doc = _load_ini(ini)
    table, final = _resolve(doc, _split_key(key))
    node = table[final]
    print(node.as_string() if hasattr(node, "as_string") else node)


def config_set(key: str, value: str, ini: str = None):
    """
        Sets a setting, preserving comments and layout, e.g.
        ``xloil config set xlOil_Python.LoadModules "['xloil.xloil_ribbon', 'mymod']"``
        ``xloil config set Addin.LogLevel "'info'"``
    """
    path, doc = _load_ini(ini)
    parts = _split_key(key)
    parsed = _parse_value(value)
    if parts[-1] in _LIST_SETTINGS and not (
            isinstance(parsed, list) and all(isinstance(x, str) for x in parsed)):
        raise ValueError(f"{parts[-1]} must be a list of names in square brackets, e.g. "
                         f"\"['xloil.xloil_ribbon', 'mymodule']\". Got: {value}")
    table, final = _resolve(doc, parts, create=True)
    table[final] = parsed
    path.write_text(_dumps(doc), encoding='utf-8')
    print(f"Set {key} = {table[final]!r} in {path}")


def _dumps(doc) -> str:
    import tomlkit
    return tomlkit.dumps(doc)


def _edit_load_modules(ini, change):
    from .features import order_modules
    path, doc = _load_ini(ini)
    table, final = _resolve(doc, ["xlOil_Python", "LoadModules"], create=True)
    before = [str(m) for m in table.get(final, [])]
    after = order_modules(change(before))
    table[final] = after
    path.write_text(_dumps(doc), encoding='utf-8')
    print(f"LoadModules = {after}\nRestart Excel to apply.")


def config_add_module(module: str, ini: str = None):
    """ Adds a module to LoadModules (kept in dependency order), e.g. your own module """
    _edit_load_modules(ini, lambda mods: mods + [module] if module not in mods else mods)


def config_remove_module(module: str, ini: str = None):
    """ Removes a module from LoadModules """
    _edit_load_modules(ini, lambda mods: [m for m in mods if m != module])


def config_add_path(folder: str, ini: str = None):
    """
        Adds a folder to xlOil's Python module search path (XLOIL_PYTHON_PATH),
        e.g. the folder holding your own modules listed in LoadModules.
    """
    folder = str(Path(folder).expanduser().resolve())
    if not Path(folder).is_dir():
        raise FileNotFoundError(f"{folder} is not a folder")
    path, doc = _load_ini(ini)
    table, final = _resolve(doc, ["xlOil_Python", "Environment", "XLOIL_PYTHON_PATH"], create=True)
    current = [p for p in str(table.get(final, "")).split(";") if p.strip()]
    if any(os.path.normcase(p) == os.path.normcase(folder) for p in current):
        print(f"{folder} is already on the search path")
        return
    table[final] = ";".join(current + [folder])
    path.write_text(_dumps(doc), encoding='utf-8')
    print(f"Added {folder} to the search path in {path}. Restart Excel to use it.")


def config_list(ini: str = None):
    """ Prints all settings in the ini file (comments omitted) """
    _, doc = _load_ini(ini)
    def walk(node, prefix):
        for k, v in node.items():
            if hasattr(v, "items"):
                walk(v, f"{prefix}{k}.")
            elif isinstance(v, list) and v and all(hasattr(x, "items") for x in v):
                for item in v:
                    walk(item, f"{prefix}{k}.")
            else:
                print(f"{prefix}{k} = {v!r}")
    walk(doc, "")


def _set_python_executable(python: str, ini: str = None):
    """ Points the ini's PYTHONEXECUTABLE at the given interpreter (venv aware) """
    exe = Path(python)
    if exe.is_dir():
        # A venv or environment directory was given
        for candidate in [exe / "Scripts" / "python.exe", exe / "python.exe"]:
            if candidate.exists():
                exe = candidate
                break
    if not exe.exists():
        raise FileNotFoundError(f"Python executable {exe} not found")
    for cfg in [exe.parent / "pyvenv.cfg", exe.parent.parent / "pyvenv.cfg"]:
        if cfg.exists():
            print(f"Virtual environment detected ({cfg}): xlOil will read its 'home' setting")
            break
    path = _ini_path(ini)
    text = path.read_text(encoding='utf-8')
    text, ok = _write_python_path_to_ini(text, _get_xloil_bin_dir(), comment_reg_keys=True,
                                         replace_paths={'PYTHONEXECUTABLE': str(exe)})
    if not ok:
        raise RuntimeError(f"Could not find PYTHONEXECUTABLE in {path}")
    path.write_text(text, encoding='utf-8')
    print(f"PYTHONEXECUTABLE set to {exe}")


_RIBBON_MODULES = {"com": "xloil.xloil_ribbon", "xlam": "xloil.xlam_ribbon"}
XLAM_NAME = "xlOil.xlam"


def _load_modules(ini=None) -> list:
    """ The LoadModules list in the settings file (parsed, so comments that mention
        module names don't count); empty if it can't be read """
    try:
        _, doc = _load_ini(ini)
        return [str(m) for m in doc["xlOil_Python"]["LoadModules"]]
    except Exception:
        return []


def _user_modules(ini=None) -> list:
    """ Your own modules in LoadModules (anything not part of xlOil), so that
        re-running setup keeps them """
    return [m for m in _load_modules(ini) if m != "xloil" and not m.startswith("xloil.")]


def _xlam_source() -> Path:
    """ The xlOil.xlam shipped in this package. (Not imported from xloil.xlam_ribbon:
        that module, like the ribbon logic it uses, only loads inside Excel.) """
    return Path(__file__).with_name("data") / XLAM_NAME


def addins_must_be_signed(version=None) -> bool:
    """
        True when Office policy "Require application add-ins to be signed by Trusted
        Publisher" is on. It blocks xlOil's unsigned COM ribbon, but not xlOil.xlam.
    """
    try:
        version = version or _find_excel_version()
    except OSError:
        version = 16
    for root in (reg.HKEY_CURRENT_USER, reg.HKEY_LOCAL_MACHINE):
        for base in (r"Software\Policies\Microsoft\Office", r"Software\Microsoft\Office"):
            with suppress(OSError):
                key = reg.OpenKey(root, rf"{base}\{version}.0\Excel\Security")
                value, _ = reg.QueryValueEx(key, "requireaddinsig")
                if value == 1:
                    return True
    return False


def ribbon(kind: str, ini: str = None):
    """
        Chooses how xlOil's ribbon tab is drawn:
          com  - the built-in COM add-in (default)
          xlam - xlOil.xlam in XLSTART: for computers whose policy requires signed add-ins
          off  - no ribbon tab
    """
    kind = kind.lower()
    if kind not in ("com", "xlam", "off"):
        raise ValueError("Choose com, xlam or off")
    target = _XL_START_PATH / XLAM_NAME
    if kind == "xlam":
        src = _xlam_source()
        if not src.exists():
            raise FileNotFoundError(f"{src} is missing: reinstall xlOil")
        _XL_START_PATH.mkdir(parents=True, exist_ok=True)
        try:
            sh.copy(src, target)
        except PermissionError:
            raise SystemExit(f"Can't replace {target}: close Excel and try again")
        _unblock(target)
        print("Installed ", target)
    elif target.exists():
        try:
            target.unlink()
        except PermissionError:
            raise SystemExit(f"Can't remove {target}: close Excel and try again")
        print("Removed ", target)
    wanted = _RIBBON_MODULES.get(kind)
    _edit_load_modules(ini, lambda mods: [m for m in mods if m not in _RIBBON_MODULES.values()]
                       + ([wanted] if wanted else []))


def _install_cmd(ini_template=None, python=None, load_modules=None, startup_script=None,
                 replace_ini=False, intellisense=False):
    install_xloil(ini_template, replace_ini=replace_ini)
    if python:
        _set_python_executable(python)
    if load_modules:
        mods = [m.strip() for m in load_modules.split(',') if m.strip()]
        config_set('xlOil_Python.LoadModules', '[' + ', '.join(f'"{m}"' for m in mods) + ']')
    if startup_script:
        config_set('xlOil_Python.StartupScript',
                   " + str(Path(startup_script).resolve()) + ")
    if intellisense:
        _intellisense_cmd('install')


def status():
    """ Prints where xlOil is installed and which Python it uses """
    print(f"xlOil Python package : {Path(__file__).parent}")
    print(f"Python executable    : {sys.executable}")
    try:
        print(f"xlOil binaries       : {_get_xloil_bin_dir()}")
    except Exception as e:
        print(f"xlOil binaries       : NOT FOUND ({e})")
    print(f"Add-in in XLSTART    : {_XLL_INSTALL_PATH} "
          f"({'present' if _XLL_INSTALL_PATH.exists() else 'missing'})")
    ini = _ini_path()
    print(f"Settings file        : {ini} ({'present' if ini.exists() else 'missing'})")
    if ini.exists():
        try:
            _, doc = _load_ini()
            py = doc.get('xlOil_Python', {})
            print(f"  LoadModules        : {py.get('LoadModules')}")
            print(f"  StartupScript      : {py.get('StartupScript')}")
            for env in py.get('Environment', []):
                if 'PYTHONEXECUTABLE' in env:
                    print(f"  PYTHONEXECUTABLE   : {env['PYTHONEXECUTABLE']}")
        except Exception as e:
            print(f"  (could not parse: {e})")
    try:
        print(f"Excel version        : {_find_excel_version()}")
    except Exception:
        print("Excel version        : not found")
    from .intellisense import default_xll_name
    isense = _XL_START_PATH / default_xll_name()
    print(f"IntelliSense add-in  : {'installed' if isense.exists() else 'not installed'}")


def _intellisense_cmd(action, target_dir=None):
    from .intellisense import install_intellisense_addin, default_xll_name
    if action == 'install':
        install_intellisense_addin(target_dir)
        try:
            _, doc = _load_ini()
            enabled = "xloil.features.intellisense" in list(doc["xlOil_Python"]["LoadModules"])
        except Exception:
            enabled = False
        if not enabled:
            print("To show descriptions, turn on IntelliSense in the ribbon's Features menu, or run:\n"
                  "  xloil config add-module xloil.features.intellisense\n"
                  "then restart Excel")
    elif action == 'remove':
        path = Path(target_dir or _XL_START_PATH) / default_xll_name()
        if path.exists():
            path.unlink()
            print(f"Removed {path}")
        else:
            print(f"{path} not found")


def _ask(question: str, default: bool, assume_yes: bool) -> bool:
    if assume_yes:
        return default
    hint = "[Y/n]" if default else "[y/N]"
    while True:
        try:
            answer = input(f"{question} {hint} ").strip().lower()
        except EOFError:
            return default
        if not answer:
            return default
        if answer in ("y", "yes"):
            return True
        if answer in ("n", "no"):
            return False


def setup(features: str = None, yes: bool = False, no_pip: bool = False, intellisense: bool = None):
    """
        Guided first-time set-up: choose features, install the libraries they
        need, install the add-in with the right LoadModules, then check everything.
        Non-interactive: ``xloil setup --features pandas,charts,qt --yes``.
    """
    import subprocess
    from .features import FEATURES, find_feature, feature_available, order_modules

    print("xlOil setup\n")
    print(f"Python: {sys.executable} ({sys.version.split()[0]})")
    home = python_home_for(sys.executable)
    if home:
        print(f"Conda installation detected ({home}): PYTHONHOME will be set automatically")
    if not _check_bitness():
        print("Fix the bitness mismatch first, then run `xloil setup` again.")
        return 1
    print()

    if features is not None:
        chosen = [find_feature(k.strip()) for k in features.split(",") if k.strip()]
    else:
        defaults = {"pandas", "charts"}
        chosen = []
        print("Choose features (Enter accepts the default):")
        for f in FEATURES:
            default = f.key in defaults and (feature_available(f) or bool(f.pip))
            installed = "" if feature_available(f) or not f.pip else f"  [will pip install {f.pip}]"
            if _ask(f"  {f.label}: {f.description}{installed}?", default, yes):
                chosen.append(f)
        print()

    missing = sorted({pkg for f in chosen if not feature_available(f) and f.pip
                      for pkg in f.pip.split()})
    if missing and not no_pip:
        print(f"Installing: {' '.join(missing)}")
        result = subprocess.run([sys.executable, "-m", "pip", "install", "--no-warn-script-location", *missing])
        if result.returncode != 0:
            print("WARNING: pip install failed; those features will not work until the "
                  "packages are installed")
        print()

    use_xlam = addins_must_be_signed()
    if use_xlam:
        print("Office policy here requires signed add-ins, which blocks xlOil's COM ribbon:\n"
              "using the xlOil.xlam ribbon instead.\n")
    modules = [_RIBBON_MODULES["xlam" if use_xlam else "com"]]
    for f in chosen:
        modules += list(f.modules)
    own = _user_modules()
    if own:
        print(f"Keeping your modules: {', '.join(own)}\n")
    modules = order_modules(modules + [m for m in own if m not in modules])

    _install_cmd(load_modules=",".join(modules))
    if use_xlam:
        ribbon("xlam")

    want_isense = intellisense
    if want_isense is None:
        want_isense = any(f.key == "intellisense" for f in chosen) and \
            _ask("Download the Excel-DNA IntelliSense add-in from GitHub?", True, yes)
    if want_isense:
        try:
            _intellisense_cmd('install')
        except Exception as e:
            print(f"WARNING: IntelliSense add-in not installed: {e}")

    print()
    problems = doctor()
    print("\nDone. Start (or restart) Excel: the 'xlOil Py' ribbon tab lets you change "
          "features later." if not problems else
          "\nSet-up finished with problems: see the messages above.")
    return problems


def doctor():
    """
        Checks the things which most often stop xlOil loading on a new computer
        and prints what to do about each problem. Returns the number of problems.
    """
    import importlib.util
    import struct
    problems = 0

    def ok(msg):
        print(f"  [ OK ] {msg}")

    def bad(msg, fix):
        nonlocal problems
        problems += 1
        print(f"  [FAIL] {msg}\n         -> {fix}")

    def warn(msg):
        print(f"  [WARN] {msg}")

    print("xlOil doctor\n")
    py_bits = struct.calcsize("P") * 8
    ver = f"{sys.version_info.major}{sys.version_info.minor}"

    exe = excel_path()
    if exe is None:
        bad("Excel not found", "Install Microsoft Excel (desktop, 2013 or later)")
    else:
        ok(f"Excel: {exe}")
        bits = excel_bitness()
        if bits is not None and bits != py_bits:
            bad(f"Excel is {bits}-bit, Python is {py_bits}-bit",
                f"Use a {bits}-bit Python with a {bits}-bit xlOil build")
        elif bits is not None:
            ok(f"Excel and Python are both {bits}-bit")

        runtime = ["msvcp140.dll", "vcruntime140.dll", "vcruntime140_1.dll"]
        system32 = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32"
        missing = [d for d in runtime if not (exe.parent / d).exists() and not (system32 / d).exists()]
        if missing:
            bad(f"Visual C++ runtime missing: {', '.join(missing)}",
                "Install the Microsoft Visual C++ Redistributable (x64): "
                "https://aka.ms/vs/17/release/vc_redist.x64.exe")
        else:
            ok("Visual C++ runtime available")

    try:
        bin_dir = _get_xloil_bin_dir()
        ok(f"xlOil binaries: {bin_dir}")
        pyd = bin_dir / f"xlOil_Python{ver}.pyd"
        if pyd.exists():
            ok(f"Extension for Python {sys.version_info.major}.{sys.version_info.minor}: {pyd.name}")
        else:
            bad(f"No xlOil extension for Python {sys.version_info.major}.{sys.version_info.minor}",
                "Install the xloil wheel built for this Python version")
        if any(Path(str(f) + ":Zone.Identifier").exists() for f in [bin_dir / ADDIN_NAME]):
            bad("xlOil binaries are marked as downloaded from the internet",
                "Run `xloil install` again (it unblocks them) or Unblock in the file Properties")
    except Exception as e:
        bad(f"xlOil binaries not found ({e})", "pip install the xloil wheel, then run `xloil install`")

    for mod, pip_name in [("numpy", "numpy"), ("win32com", "pywin32"), ("tomlkit", "tomlkit")]:
        if importlib.util.find_spec(mod) is None:
            bad(f"Python package '{mod}' missing", f"pip install {pip_name}")
        else:
            ok(f"Python package '{mod}' available")

    home = python_home_for(sys.executable)
    ini = _ini_path()
    if not _XLL_INSTALL_PATH.exists():
        warn(f"xlOil.xll not in XLSTART: run `xloil install` (or load an add-in made by `xloil create`)")
    else:
        ok(f"Add-in installed: {_XLL_INSTALL_PATH}")
        _, blocked = _XLL_INSTALL_PATH, Path(str(_XLL_INSTALL_PATH) + ":Zone.Identifier").exists()
        if blocked:
            bad("Installed add-in is blocked (Mark of the Web)", "Run `xloil install` again")

    if ini.exists():
        text = ini.read_text(encoding='utf-8')
        m = re.search(r"^\s*PYTHONEXECUTABLE\s*=\s*'''(.*?)'''", text, flags=re.M)
        configured = m.group(1).replace('\\\\', '\\') if m else ""
        if not configured:
            bad("PYTHONEXECUTABLE not set in settings", "Run `xloil install`")
        elif not Path(configured).exists():
            bad(f"PYTHONEXECUTABLE points to a missing file: {configured}",
                "Run `xloil config python <path to python.exe>`")
        else:
            ok(f"Settings use Python: {configured}")
            needed = python_home_for(configured)
            has_home = re.search(r"^\s*PYTHONHOME\s*=", text, flags=re.M) is not None
            if needed and not has_home:
                bad("Conda Python without PYTHONHOME: Python will fail to start in Excel",
                    f"Run `xloil config python \"{configured}\"` to add it")
            elif needed:
                ok(f"Conda PYTHONHOME set")
        mods = _load_modules(ini)   # not `text`: the ini's comments mention module names too
        signed_only = addins_must_be_signed()
        xlam_installed = (_XL_START_PATH / XLAM_NAME).exists()
        if _RIBBON_MODULES["com"] in mods and signed_only:
            bad("Office policy requires signed add-ins, so the COM ribbon can't load "
                "(you'd see a 'COM addin connect' error)", "Run `xloil ribbon xlam`")
        elif _RIBBON_MODULES["xlam"] in mods and not xlam_installed:
            bad("LoadModules has xloil.xlam_ribbon but xlOil.xlam isn't in XLSTART",
                "Run `xloil ribbon xlam`")
        elif _RIBBON_MODULES["xlam"] in mods:
            ok("Ribbon: xlOil.xlam")
        elif _RIBBON_MODULES["com"] in mods:
            ok("Ribbon: COM add-in")
    else:
        warn(f"No settings file at {ini}: run `xloil install`")

    if (Path(sys.prefix) / "conda-meta").exists():
        if importlib.util.find_spec("qtpy") is not None or importlib.util.find_spec("PySide6") is not None:
            try:
                import PySide6
                if (Path(PySide6.__file__).parent / "Qt6Core.dll").exists():
                    ok("Conda env with pip PySide6: xlOil avoids the ICU DLL clash automatically")
                else:
                    warn("conda-forge PySide6 on Windows has no Qt WebEngine, so interactive charts "
                         "fall back to pictures. For interactive charts: "
                         "conda remove pyside6 && pip install PySide6")
            except ImportError:
                pass

    print(f"\n{problems} problem(s) found" if problems else "\nNo problems found")
    return problems


def main():
    parser = ArgumentParser(prog='xlOil', 
                            description='Excel/Python integration')
    
    subparsers = parser.add_subparsers(help='Commands to control local xlOil distribution')

    p = subparsers.add_parser('install', 
                              help='Installs xloil by copying the Excel addin to the XLSTART directory '
                                   'and sets the correct paths in xloil.ini')
    p.add_argument("ini_template", nargs='?', type=str, default=None,
                   help='Path to xloil.ini file to modify and install, if not present the default file is used')
    p.add_argument("--python", type=str, default=None,
                   help='Python executable, environment or venv directory xlOil should use '
                        '(default: the one running this command)')
    p.add_argument("--load-modules", dest="load_modules", type=str, default=None,
                   help='Comma-separated modules to load on start-up, e.g. xloil.xloil_ribbon,mymodule')
    p.add_argument("--startup-script", dest="startup_script", type=str, default=None,
                   help='A python script to run when xlOil starts, before LoadModules')
    p.add_argument("--replace-ini", dest="replace_ini", action='store_true',
                   help='Overwrite an existing settings file rather than keeping it')
    p.add_argument("--intellisense", action='store_true',
                   help='Also download and install the Excel-DNA IntelliSense add-in')
    p.set_defaults(func=_install_cmd)

    p = subparsers.add_parser('status', help='Shows where xlOil is installed and how it is configured')
    p.set_defaults(func=status)

    p = subparsers.add_parser('doctor', help='Checks for common problems which stop xlOil loading')
    p.set_defaults(func=doctor)

    p = subparsers.add_parser('setup', help='Guided first-time set-up: pick features, install what they '
                                            'need, install the add-in and check it')
    p.add_argument("--features", default=None,
                   help='Comma-separated features, skipping the questions: '
                        'pandas,polars,charts,qt,intellisense,cache')
    p.add_argument("--yes", "-y", action='store_true', help='Accept the defaults without asking')
    p.add_argument("--no-pip", dest="no_pip", action='store_true',
                   help="Don't pip install the libraries features need")
    p.set_defaults(func=setup)

    p = subparsers.add_parser('ribbon', help="Chooses how the 'xlOil Py' ribbon tab is drawn: com (default), "
                                             "xlam (for PCs whose policy requires signed add-ins) or off")
    p.add_argument("kind", choices=["com", "xlam", "off"])
    p.set_defaults(func=ribbon)

    p = subparsers.add_parser('config', help='Reads or changes settings in xlOil.ini')
    config_sub = p.add_subparsers()
    c = config_sub.add_parser('get', help='Prints a setting, e.g. xlOil_Python.LoadModules')
    c.add_argument("key")
    c.add_argument("--ini", default=None, help='Settings file (default %%APPDATA%%/xlOil/xlOil.ini)')
    c.set_defaults(func=config_get)
    c = config_sub.add_parser('set', help='Sets a setting. The value is parsed as TOML, so quote strings')
    c.add_argument("key")
    c.add_argument("value")
    c.add_argument("--ini", default=None)
    c.set_defaults(func=config_set)
    c = config_sub.add_parser('list', help='Prints all settings')
    c.add_argument("--ini", default=None)
    c.set_defaults(func=config_list)
    c = config_sub.add_parser('python', help='Sets the Python executable / venv xlOil uses')
    c.add_argument("python")
    c.add_argument("--ini", default=None)
    c.set_defaults(func=_set_python_executable)
    c = config_sub.add_parser('add-module', help='Adds a module to LoadModules, e.g. your own module')
    c.add_argument("module")
    c.add_argument("--ini", default=None)
    c.set_defaults(func=config_add_module)
    c = config_sub.add_parser('remove-module', help='Removes a module from LoadModules')
    c.add_argument("module")
    c.add_argument("--ini", default=None)
    c.set_defaults(func=config_remove_module)
    c = config_sub.add_parser('add-path', help='Adds a folder of your own Python modules to the search path')
    c.add_argument("folder")
    c.add_argument("--ini", default=None)
    c.set_defaults(func=config_add_path)

    p = subparsers.add_parser('intellisense',
                              help='Installs or removes the Excel-DNA IntelliSense add-in '
                                   '(in-sheet function tooltips)')
    p.add_argument("action", choices=['install', 'remove'])
    p.add_argument("--target-dir", dest="target_dir", default=None,
                   help='Install directory (default: Excel XLSTART folder)')
    p.set_defaults(func=_intellisense_cmd)
    
    p = subparsers.add_parser('create', 
                              help='Creates an XLL addin to simplify distribution of Excel funcs to other xlOil users')
    p.add_argument("filename", type=str, 
                   help='Name of the XLL file to create')
    p.set_defaults(func=_create_addin)

    p = subparsers.add_parser('package', 
                              help='Uses PyInstaller to create a packaged xloil installer. Any arguments '
                                   'which xlOil does not parse will be passed directly to PyInstaller. '
                                   'Invoke this function from a minimal python distribution, otherwise the '
                                   'resulting package may be large. See docs for more details.')
    p.add_argument("ini_template", type=str, 
                   help='The xlOil ini file to package and distribute')
    p.add_argument("--makespec", action='store_true', 
                   help='If present, creates the PyInstaller spec file then exits. This allows more precise '
                        'tweaking of the PyInstaller config. See PyInstaller docs on pyi-makespec')
    p.set_defaults(func=_package_pyinstaller)

    p = subparsers.add_parser('clean', 
                              help='Uninstalls xlOil and removes its package using pip')
    p.set_defaults(func=_clean_xloil)

    p = subparsers.add_parser('remove', 
                              help='Removes the xlOil addin from Excel. Leaves xlOil.ini in place')
    p.set_defaults(func=_remove_xloil)
    
    p = subparsers.add_parser('uninstall', 
                              help='Equivalent to the remove command')
    p.set_defaults(func=_remove_xloil)
    
    args, more_args = parser.parse_known_args()
    
    func_args = vars(args)
    func = func_args.pop("func", None)
    if func is None:
        parser.print_help()
        return
    if len(more_args) > 0:
        func(**vars(args), extra_args=more_args)
    else:
        func(**vars(args))
        

if __name__ == '__main__':
    main()