"""Launcher kept at the repository root so ``python main.py`` still works.

The application itself lives in ``src/eventplayback``; this shim only makes the
package importable from a plain source checkout.
"""

from __future__ import annotations

import sys
from pathlib import Path

_SRC = Path(__file__).resolve().parent / "src"
if _SRC.is_dir() and str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from eventplayback.__main__ import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
