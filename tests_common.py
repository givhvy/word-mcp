"""Shared test setup: UTF-8 output on any console, and one suite at a time per Word instance."""

import msvcrt
import os
import sys
import tempfile

_lock_file = None


def setup() -> None:
    # Document text can be any language; a cp1252 console must not crash the run.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    # All suites automate the same Word; two at once make each other's documents disappear
    # ("Object has been deleted"). Hold a lock for the whole run.
    global _lock_file
    path = os.path.join(tempfile.gettempdir(), "word-mcp-tests.lock")
    _lock_file = open(path, "a+")
    try:
        msvcrt.locking(_lock_file.fileno(), msvcrt.LK_NBLCK, 1)
    except OSError:
        sys.exit("another word-mcp test suite is already running against Word; run the suites one at a time")


def doc_diff(before: set, after: set) -> str:
    extra, missing = sorted(after - before), sorted(before - after)
    return f"left open: {extra or 'none'}; closed: {missing or 'none'}"
