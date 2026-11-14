"""
    Packages a Release x64 build into distributables, for CI (GitHub Actions).
    A lighter alternative to stage.py, which also builds Win32 and the
    Sphinx/Doxygen docs.

    Usage (from the repository root, after `msbuild xlOil.sln /p:Configuration=Release /p:Platform=x64`):

        python tools/ci_package.py --version 0.22.1.post19 --python C:/py310/python.exe --python C:/py312/python.exe

    Produces in build/dist:
        xlOil-<version>-x64-bin.zip    the add-in binaries, ini and scripts
        xloil-<version>-cpXY-cpXY-win_amd64.whl   one wheel per --python given
"""
import argparse
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BUILD = ROOT / "build"
STAGING = BUILD / "staging"
DIST = BUILD / "dist"
ARCH = "x64"

CORE_FILES = ["xlOil.xll", "xlOil.dll", "xlOil_Python.dll", "xlOil_SQL.dll", "xlOil_Utils.dll"]
SCRIPTS = ["xlOil_Install.ps1", "xlOil_Remove.ps1", "xlOil_NewAddin.ps1"]


def python_version(exe: str) -> str:
    out = subprocess.run([exe, "-c", "import sys; print(f'{sys.version_info[0]}.{sys.version_info[1]}')"],
                         capture_output=True, text=True, check=True)
    return out.stdout.strip()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--version", required=True, help="Package version, e.g. 0.22.1.post19")
    parser.add_argument("--python", action="append", default=[], help="Python executable to build a wheel for")
    parser.add_argument("--bin-dir", default=str(BUILD / ARCH / "Release"))
    args = parser.parse_args()

    bin_dir = Path(args.bin_dir)
    shutil.rmtree(STAGING, ignore_errors=True)
    DIST.mkdir(parents=True, exist_ok=True)
    arch_dir = STAGING / ARCH
    arch_dir.mkdir(parents=True)

    # Binaries
    pyds = sorted(bin_dir.glob("xlOil_Python3*.pyd"))
    if not pyds:
        sys.exit(f"No xlOil_Python3*.pyd found in {bin_dir}")
    for name in CORE_FILES:
        shutil.copy2(bin_dir / name, arch_dir)
    for pyd in pyds:
        shutil.copy2(pyd, arch_dir)
    for name in SCRIPTS:
        shutil.copy2(ROOT / "tools" / name, arch_dir)
    shutil.copy2(ROOT / "config" / "xlOil.ini", STAGING)
    shutil.copy2(ROOT / "Version.txt", STAGING)
    built_versions = {p.stem.replace("xlOil_Python", "") for p in pyds}
    print("Built python extensions:", ", ".join(sorted(built_versions)))

    # Binary zip
    zip_path = DIST / f"xlOil-{args.version}-{ARCH}-bin.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(arch_dir.iterdir()):
            z.write(f, f"{ARCH}/{f.name}")
        z.write(STAGING / "xlOil.ini", "xlOil.ini")
        z.write(STAGING / "Version.txt", "Version.txt")
    print("Wrote", zip_path)

    # Python package source, as stage.py lays it out
    pkg = STAGING / "pypackage"
    shutil.copytree(ROOT / "libs" / "xlOil_Python" / "Package", pkg,
                    ignore=shutil.ignore_patterns("tests", "__pycache__", "*.pyproj", "test_*.py",
                                                  "build", "dist", "*.egg-info"))
    shutil.copy2(ROOT / "LICENSE", pkg)
    shutil.copy2(ROOT / "README.md", pkg)

    for exe in args.python:
        ver = python_version(exe)
        if ver.replace(".", "") not in built_versions:
            print(f"Skipping Python {ver}: no matching extension was built")
            continue
        subprocess.run([exe, "-m", "pip", "install", "--quiet", "--upgrade", "setuptools", "wheel"], check=True)
        shutil.rmtree(pkg / "build", ignore_errors=True)
        cmd = [exe, "setup.py", "bdist_wheel", "--arch", ARCH, "--pyver", ver,
               "--version", args.version, "--plat-name", "win_amd64"]
        print("Running:", " ".join(cmd))
        subprocess.run(cmd, cwd=pkg, check=True)

    for whl in (pkg / "dist").glob("*.whl"):
        shutil.copy2(whl, DIST)
        print("Wrote", DIST / whl.name)


if __name__ == "__main__":
    main()
