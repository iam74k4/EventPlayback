"""Reading and writing macro files."""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path

from ..core.macro import Macro

__all__ = ["MAX_FILE_BYTES", "MacroFileError", "load_macro", "save_macro"]

#: Refuse to parse anything larger than this. A macro of a few hours is well
#: under a megabyte, so a bigger file is a mistake, not a recording.
MAX_FILE_BYTES = 64 * 1024 * 1024


class MacroFileError(Exception):
    """Raised when a macro cannot be read or written."""


def save_macro(macro: Macro, path: str | os.PathLike[str]) -> Path:
    """Write ``macro`` to ``path`` atomically.

    The file is written to a temporary sibling and then moved into place, so an
    interrupted save cannot leave a half-written macro behind.
    """
    target = Path(path).expanduser()
    if target.suffix.lower() != ".json":
        target = target.with_suffix(".json")

    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(macro.to_dict(), ensure_ascii=False, indent=2)
        handle, temp_name = tempfile.mkstemp(
            dir=str(target.parent), prefix=f".{target.stem}-", suffix=".tmp"
        )
        try:
            with os.fdopen(handle, "w", encoding="utf-8") as stream:
                stream.write(payload)
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temp_name, target)
        except BaseException:
            try:
                os.unlink(temp_name)
            except OSError:
                pass
            raise
    except PermissionError as exc:
        raise MacroFileError("Permission denied, or the file is in use") from exc
    except OSError as exc:
        raise MacroFileError(f"Save error: {exc}") from exc
    return target


def load_macro(path: str | os.PathLike[str]) -> Macro:
    """Read a macro from ``path``."""
    source = Path(path).expanduser()
    try:
        if not source.is_file():
            raise MacroFileError("File not found")
        size = source.stat().st_size
        if size > MAX_FILE_BYTES:
            raise MacroFileError(f"File is too large ({size / 1024 / 1024:.0f} MB)")
        text = source.read_text(encoding="utf-8")
    except MacroFileError:
        raise
    except UnicodeDecodeError as exc:
        raise MacroFileError("File is not valid UTF-8 text") from exc
    except PermissionError as exc:
        raise MacroFileError("Permission denied") from exc
    except OSError as exc:
        raise MacroFileError(f"Load error: {exc}") from exc

    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise MacroFileError(f"JSON format error: {exc}") from exc

    try:
        macro = Macro.from_dict(data)
    except ValueError as exc:
        raise MacroFileError(f"Data format error: {exc}") from exc

    if not macro.name or macro.name == "Untitled":
        macro.name = source.stem
    return macro
