"""COMP6009 Lab 2 launcher for the improved CoinTex genetic agent.

The original tutorial in this folder duplicates an older 24-level game.  The
improved implementation lives in ``PlayerGA/main.py`` and reuses the current
CoinTex engine, so this launcher preserves the lab's requested command without
maintaining a second copy of the game.

Run from this folder with::

    python main_updated.py
"""

from pathlib import Path
import sys


PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import runpy


if __name__ == "__main__":
    runpy.run_path(str(PROJECT_ROOT / "main_updated.py"), run_name="__main__")
