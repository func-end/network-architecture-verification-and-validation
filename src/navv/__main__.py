#!/usr/bin/env python3

"""NAVV CLI entrypoint.

This module is the earliest safe place to sanitize argv before Click parses it.
In frozen/PyInstaller builds (especially with multiprocessing), helper processes
can be spawned with Python-style flags (e.g., -B). Click does not recognize
these and will error unless we remove them first.
"""

from __future__ import annotations

import multiprocessing
import sys


def _sanitize_argv(argv: list[str]) -> list[str]:
    """Remove Python interpreter flags that can appear in frozen relaunches."""
    # Most common offender observed in the field.
    python_flags = {"-B"}
    return [a for a in argv if a not in python_flags]


def main() -> None:
    # Ensure spawned subprocesses on Windows/macOS frozen builds behave.
    multiprocessing.freeze_support()

    # Sanitize argv *before* importing anything that invokes Click.
    sys.argv = _sanitize_argv(sys.argv)

    from navv.network_analysis import main

    main()


if __name__ == "__main__":
    main()
