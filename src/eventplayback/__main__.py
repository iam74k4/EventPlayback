"""Entry point: ``python -m eventplayback``."""

from __future__ import annotations

import logging

from .platform_support import enable_dpi_awareness

__all__ = ["main"]


def _configure_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def main() -> int:
    _configure_logging()
    # Must happen before the toolkit creates any window.
    enable_dpi_awareness()

    from .ui.app import App

    App().mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
