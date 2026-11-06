"""Run the stdio end-to-end scenario (real server, simulated Outlook)."""

import asyncio
import sys

import pytest

from tests.e2e.run_e2e import run_all


@pytest.mark.skipif(sys.version_info < (3, 10), reason="needs 3.10+")
def test_end_to_end_over_stdio():
    results = asyncio.run(run_all())
    failed = [f"{n}: {d}" for s, n, d in results if s != "PASS"]
    assert not failed, "\n".join(failed)
    assert len(results) >= 40
