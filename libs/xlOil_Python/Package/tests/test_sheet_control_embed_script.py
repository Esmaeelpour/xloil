"""Runs libs/xlOil_Python/tests_py/test_sheet_control_embed.py (sheet controls embedded through their toolkit,
clipping to the scrolling pane, Excel's quit, the Qt binding choice) in its own stub-mode Python, as CI's test job."""
import os
import subprocess
import sys
from pathlib import Path

PACKAGE = Path(__file__).resolve().parents[1]
SCRIPT = PACKAGE.parent / "tests_py" / "test_sheet_control_embed.py"


def test_sheet_control_embed():
    env = dict(os.environ, XLOIL_STUBS="1", PYTHONPATH=str(PACKAGE))
    env.pop("QT_API", None)
    run = subprocess.run([sys.executable, str(SCRIPT)], env=env, capture_output=True, text=True, timeout=300)
    assert run.returncode == 0 and "ALL PASS" in run.stdout, run.stdout + run.stderr
