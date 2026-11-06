"""DASL dates are UTC; extra allowed attachment folders."""

import datetime as dt
import os

import pytest

from outlook_mcp.errors import OutlookError
from outlook_mcp.utils.formatting import to_dasl_date
from outlook_mcp.utils.paths import validate_attachment_path


def test_to_dasl_date_converts_local_to_utc():
    naive = dt.datetime(2026, 7, 1, 12, 0)
    expected = naive.astimezone().astimezone(dt.timezone.utc).strftime("%Y-%m-%d %H:%M")
    assert to_dasl_date(naive) == expected


def test_to_dasl_date_with_explicit_offset():
    aware = dt.datetime(2026, 7, 1, 12, 0, tzinfo=dt.timezone(dt.timedelta(hours=3)))
    assert to_dasl_date(aware) == "2026-07-01 09:00"


def test_allowed_paths_env_extends_sandbox(monkeypatch, tmp_path):
    profile = tmp_path / "profile"
    profile.mkdir()
    drive = tmp_path / "drive"
    drive.mkdir()
    f = drive / "doc.pdf"
    f.write_text("x")
    monkeypatch.setenv("USERPROFILE", str(profile))
    with pytest.raises(OutlookError, match="OUTLOOK_MCP_ALLOWED_PATHS"):
        validate_attachment_path(str(f))
    monkeypatch.setenv("OUTLOOK_MCP_ALLOWED_PATHS", os.pathsep.join(["", str(drive)]))
    assert validate_attachment_path(str(f)) == str(f.resolve())
