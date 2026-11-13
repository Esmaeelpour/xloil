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
        user_access, = reg.QueryValueEx(user_key, "AccessVBOM")
    with suppress(OSError):
        lm_key =  reg.OpenKey(reg.HKEY_LOCAL_MACHINE, _excel_regpath(version) + r"\Security")
        lm_access, = reg.QueryValueEx(lm_key, "AccessVBOM")

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

def _get_python_paths():
    """
        Returns the paths to be set in the xlOil.ini file (the appropriate stubs
        must already exist).
    """
    return { 
        'PYTHONEXECUTABLE': sys.executable
    }

def _write_python_path_to_ini(ini_txt, bin_dir:str, comment_reg_keys:bool, replace_paths=None):

    if replace_paths is None:
        replace_paths = _get_python_paths()

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
    
    # Comment out the now usused code to get the python paths from the registry
    # Don't error if this fails as it's not critical
    if comment_reg_keys:
        
        for key in ["xlOil_PythonRegKey"]:
            replace(rf'^(\s*{key}\s*=.*)', r'#\g<1>')

    return ini_txt, fails == 0
    
   
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

    # Ensure XLSTART dir really exists
    with suppress(FileExistsError):
        os.mkdir(_XL_START_PATH)

    bin_dir = _get_xloil_bin_dir()

    # Copy the XLL
    sh.copy(bin_dir / ADDIN_NAME, _XLL_INSTALL_PATH)
    print("Installed ", _XLL_INSTALL_PATH)
    
    # Copy the ini file to APPDATA, avoiding overwriting any existing ini
    if not replace_ini and ini_path.exists():
        print("Found existing settings file at \n", ini_path)
    else:
        with suppress(FileExistsError):
            ini_path.parent.mkdir()
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

    excel_version = _find_excel_version()
    
    # Ensure no xlOil addins are in the registry
    _remove_addin(excel_version, _XLL_INSTALL_PATH)
    try:
        os.remove(_XLL_INSTALL_PATH)
    except FileNotFoundError:
        ...

def _clean_xloil():

    _remove_xloil()

    ini_path = Path(APP_DATA_DIR) / INIFILE_NAME
    try:
        os.remove(os.path.join(APP_DATA_DIR, INIFILE_NAME))
        os.remove(os.path.join(APP_DATA_DIR, "xlOil.log"))
    except FileNotFoundError:
        ...
    import subprocess
    import sys

    subprocess.Popen(f"{sys.executable} -m pip uninstall --yes xloil", shell=True)


def _create_addin(filename):

    basename = Path(os.path.splitext(filename)[0])

    xll_path = basename.with_suffix(".xll")
    ini_path = basename.with_suffix(".ini")

    bin_dir = _get_xloil_bin_dir()
    print("xlOil binaries found at:", str(bin_dir))

    sh.copy(bin_dir / ADDIN_NAME,    xll_path)
    sh.copy(bin_dir / INIFILE_NAME,  ini_path)
    
    print("New addin created at:", xll_path)

    # Edit ini file
    ini_txt = ini_path.read_text(encoding='utf-8')
    
    # Assume we want the xlOil_Python plugin as we're running a python script
    ini_txt, count = re.subn(r'^(\s*Plugins\s*=).*', r'\g<1>["xlOil_Python"]', ini_txt, flags=re.M)

    # Assume we want the python paths set to the distribution running this script
    ini_txt, success = _write_python_path_to_ini(ini_txt, bin_dir, 
                                                 comment_reg_keys=True)
    
    ini_path.write_text(ini_txt)

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


def config_get(key: str, ini: str = None):
    """ Prints the value of a setting, e.g. ``xloil config get xlOil_Python.LoadModules`` """
    _, doc = _load_ini(ini)
    node = doc
    for part in _split_key(key):
        node = node[part]
    print(node.as_string() if hasattr(node, "as_string") else node)


def config_set(key: str, value: str, ini: str = None):
    """
        Sets a setting, preserving comments and layout, e.g.
        ``xloil config set xlOil_Python.LoadModules "['xloil.xloil_ribbon', 'mymod']"``
    """
    import tomlkit
    path, doc = _load_ini(ini)
    parts = _split_key(key)
    node = doc
    for part in parts[:-1]:
        if part not in node:
            node[part] = tomlkit.table()
        node = node[part]
    node[parts[-1]] = _parse_value(value)
    path.write_text(tomlkit.dumps(doc), encoding='utf-8')
    print(f"Set {key} = {node[parts[-1]]!r} in {path}")


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
        print("Enable descriptions by adding 'xloil.intellisense' to LoadModules or calling "
              "xloil.intellisense.enable() in a startup script, then restart Excel")
    elif action == 'remove':
        path = Path(target_dir or _XL_START_PATH) / default_xll_name()
        if path.exists():
            path.unlink()
            print(f"Removed {path}")
        else:
            print(f"{path} not found")


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