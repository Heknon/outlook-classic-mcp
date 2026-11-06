"""Start the server through scripts/launch-mcp.cmd exactly as the plugin
does (cmd.exe /d /c ... run <data dir>) and complete an MCP session.

Windows only. Used by CI with the simulated Outlook from tests/e2e/fakewin
on PYTHONPATH (uv and venv Pythons honour PYTHONPATH, so the stand-ins
shadow the real pywin32).

    python tests/plugin/check_launcher.py <data dir> [max_start_seconds]
"""

from __future__ import annotations

import asyncio
import os
import sys
import time
from pathlib import Path

from mcp import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client

ROOT = Path(__file__).resolve().parents[2]


async def main(data_dir: str, max_start: float) -> int:
    env = dict(os.environ)
    env["PYTHONPATH"] = str(ROOT / "tests" / "e2e" / "fakewin")
    params = StdioServerParameters(
        command="cmd.exe",
        args=["/d", "/c", str(ROOT / "scripts" / "launch-mcp.cmd"), "run", data_dir],
        env=env,
    )
    t0 = time.monotonic()
    async with stdio_client(params) as (r, w):
        async with ClientSession(r, w) as s:
            await s.initialize()
            started = time.monotonic() - t0
            tools = (await s.list_tools()).tools
            res = await s.call_tool("outlook_whoami", {})
            text = res.content[0].text
    print(f"started in {started:.1f}s, {len(tools)} tools, whoami ok={'Ori Harel' in text}")
    assert len(tools) >= 37, len(tools)
    assert "Ori Harel" in text, text[:200]
    assert started < max_start, f"startup took {started:.1f}s (MCP_TIMEOUT default is 30s)"
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main(sys.argv[1], float(sys.argv[2]) if len(sys.argv) > 2 else 30.0)))
