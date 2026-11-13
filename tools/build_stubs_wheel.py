"""
    Builds the pure-Python ``xloil-stubs`` wheel: the xloil Python package without
    any compiled parts, always running in stub mode. Install it on Linux/macOS
    (or in Windows CI without Excel) to import and unit test code that uses xloil.

        python tools/build_stubs_wheel.py --version 0.22.1.post2

    The wheel is written to build/dist. Do not install it alongside the full
    ``xloil`` package: both provide the ``xloil`` module.
"""
import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

SETUP = '''
from setuptools import setup, find_packages
setup(
    name="xloil-stubs",
    version="{version}",
    description="Stub (testing) build of xlOil: import and unit test xlOil code without Excel",
    long_description=open("README.md", encoding="utf-8").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/Esmaeelpour/xloil",
    license="Apache",
    packages=find_packages(),
    package_data={{"xloil": ["_STUBS_ONLY"]}},
    python_requires=">=3.8",
    install_requires=["numpy>=1.18", "tomlkit"],
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
        "Topic :: Software Development :: Testing",
    ],
)
'''

README = '''# xloil-stubs

Pure-Python stub build of [xlOil](https://github.com/Esmaeelpour/xloil) for unit
testing. `import xloil` works on any platform without Excel: decorators record
what they declare and nothing calls Excel. See `xloil.testing` for helpers:

```python
import xloil.testing as xt
from mypackage import my_func

def test_my_func():
    assert xt.call(my_func, 1, 2) == 3
    assert xt.spec(my_func).name == "my_func"
```

Do not install alongside the full `xloil` package.
'''


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--version", required=True)
    parser.add_argument("--out", default=str(ROOT / "build" / "dist"))
    args = parser.parse_args()

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        shutil.copytree(ROOT / "libs" / "xlOil_Python" / "Package" / "xloil", tmp / "xloil",
                        ignore=shutil.ignore_patterns("__pycache__", "_pyinstaller"))
        (tmp / "xloil" / "_STUBS_ONLY").write_text("This distribution always runs xloil in stub mode\n")
        (tmp / "setup.py").write_text(SETUP.format(version=args.version))
        (tmp / "README.md").write_text(README)
        shutil.copy2(ROOT / "LICENSE", tmp / "LICENSE")
        out = Path(args.out)
        out.mkdir(parents=True, exist_ok=True)
        subprocess.run([sys.executable, "-m", "pip", "wheel", "--no-deps", "--no-build-isolation",
                        "-w", str(out), str(tmp)], check=True)
    print("Wrote stubs wheel to", out)


if __name__ == "__main__":
    main()
