"""Shared test setup.

The client modules import ``pythoncom`` (pywin32), which only exists on
Windows. Stubbing it lets the COM-facing logic be tested with mocks on
any platform; real COM behaviour still has to be checked on Windows.
"""

from __future__ import annotations

import sys
import types

import pytest

if sys.platform != "win32":
    try:
        import pythoncom  # noqa: F401
    except ImportError:
        stub = types.ModuleType("pythoncom")

        class com_error(Exception):  # noqa: N801 - mirrors pywin32's name
            pass

        stub.com_error = com_error
        sys.modules["pythoncom"] = stub


@pytest.fixture(autouse=True)
def _isolate_environment(monkeypatch, tmp_path):
    """Keep tests away from the real registry, signatures and env switches."""
    import outlook_mcp.utils.signatures as sigs

    monkeypatch.setattr(sigs, "_winreg", lambda: None)
    monkeypatch.setenv("OUTLOOK_MCP_SIGNATURES_DIR", str(tmp_path / "no-signatures"))
    for var in ("OUTLOOK_MCP_DRAFT_ONLY", "OUTLOOK_MCP_ALLOWED_PATHS", "OUTLOOK_MCP_ALLOW_ANY_PATH"):
        monkeypatch.delenv(var, raising=False)
