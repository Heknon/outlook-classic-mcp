"""Draft-only is the default; sending must be enabled explicitly."""

import pytest

from outlook_mcp.utils.policy import draft_only


@pytest.mark.parametrize(
    "env,expected",
    [
        ({}, True),
        ({"OUTLOOK_MCP_ALLOW_SEND": "1"}, False),
        ({"OUTLOOK_MCP_ALLOW_SEND": "no"}, True),
        ({"OUTLOOK_MCP_DRAFT_ONLY": "0"}, False),
        ({"OUTLOOK_MCP_DRAFT_ONLY": "1", "OUTLOOK_MCP_ALLOW_SEND": "1"}, True),
    ],
)
def test_draft_only_default_and_overrides(monkeypatch, env, expected):
    monkeypatch.delenv("OUTLOOK_MCP_ALLOW_SEND", raising=False)
    monkeypatch.delenv("OUTLOOK_MCP_DRAFT_ONLY", raising=False)
    for k, v in env.items():
        monkeypatch.setenv(k, v)
    assert draft_only() is expected


def test_reply_is_a_draft_by_default(monkeypatch):
    from outlook_mcp.client import mail as mail_mod
    from tests._fakes import FakeMail, FakeNamespace, FakeOriginal

    monkeypatch.delenv("OUTLOOK_MCP_ALLOW_SEND", raising=False)
    reply = FakeMail(html="<html><body>q</body></html>")
    reply.Recipients.Add("a@x.com")
    out = mail_mod.reply_mail(None, FakeNamespace({"O": FakeOriginal(reply)}), entry_id="O", body="hi")
    assert out["status"] == "saved_to_drafts" and reply.sent == 0
