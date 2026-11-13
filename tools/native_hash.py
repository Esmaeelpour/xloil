"""
    Fingerprint of everything that goes into xlOil's compiled binaries, taken
    from git (so it is the same on any checkout and any line-ending setting).

    Python-only changes (libs/xlOil_Python/Package), docs, tools and CI files
    don't change it. CI uses it to reuse the binaries of the last verified
    release (tools/native-binaries.json) instead of compiling fresh ones when
    nothing native changed. Fresh unsigned builds get a new file fingerprint,
    which Windows Defender's machine-learning check sometimes flags.

        python tools/native_hash.py [<git revision>]      # prints the fingerprint
        python tools/native_hash.py --check               # exit 0 if HEAD matches native-binaries.json
"""
import hashlib
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
NATIVE_PATHS = ["src", "include", "external", "libs", "xlOil.sln"]
EXCLUDE_PREFIXES = ("libs/xlOil_Python/Package/",)


def native_hash(revision: str = "HEAD") -> str:
    listing = subprocess.run(
        ["git", "ls-tree", "-r", "--full-tree", revision, "--", *NATIVE_PATHS],
        cwd=ROOT, capture_output=True, text=True, check=True).stdout
    lines = [line for line in listing.splitlines()
             if not line.split("\t", 1)[-1].startswith(EXCLUDE_PREFIXES)]
    return hashlib.sha256("\n".join(sorted(lines)).encode("utf-8")).hexdigest()


def main():
    args = sys.argv[1:]
    if args and args[0] == "--check":
        recorded = json.loads((ROOT / "tools" / "native-binaries.json").read_text(encoding="utf-8"))
        current = native_hash("HEAD")
        if current == recorded["source_hash"]:
            print(recorded["release"])
            return 0
        print(f"native sources changed ({current[:12]} != {recorded['source_hash'][:12]})", file=sys.stderr)
        return 1
    print(native_hash(args[0] if args else "HEAD"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
