"""Tests for send / reply / forward / send_draft against fake COM objects."""

import pytest

from outlook_mcp.client import mail as mail_mod
from outlook_mcp.errors import OutlookError
from tests._fakes import FakeMail, FakeNamespace, FakeOriginal, fake_outlook_creating

QUOTED = (
    "<html><head><style>p.MsoNormal{font-family:Aptos}</style></head>"
    "<body lang=EN-US><div class=WordSection1><p class=MsoNormal>&nbsp;</p>"
    "<div><p><b>From:</b> Ann</p></div><p class=MsoNormal>Original</p></div></body></html>"
)


def _reply_setup(html=QUOTED, body_format=None, recipients=("ann@x.com",)):
    reply = FakeMail(html=html, body_format=body_format, entry_id="R1")
    if body_format == 1:
        reply.Body = "From: Ann\nOriginal"
    for r in recipients:
        reply.Recipients.Add(r)
    original = FakeOriginal(reply)
    ns = FakeNamespace({"ORIG": original})
    return reply, original, ns


def test_reply_save_only_keeps_thread_and_saves_draft():
    reply, original, ns = _reply_setup()
    out = mail_mod.reply_mail(None, ns, entry_id="ORIG", body="Thanks!\nOri", save_only=True)

    assert out["status"] == "saved_to_drafts"
    assert out["entry_id"] == "R1"
    assert out["in_reply_to"] == "ORIG"
    assert reply.saved == 1 and reply.sent == 0
    assert original.calls == ["Reply"]  # a real reply, so it stays threaded
    html = reply.HTMLBody
    assert html.index("Thanks!") > html.index("<div class=WordSection1>")
    assert html.index("Thanks!") < html.index("From:")
    assert "p.MsoNormal{font-family:Aptos}" in html  # head styles kept


def test_reply_sends_by_default_and_reply_all():
    reply, original, ns = _reply_setup()
    out = mail_mod.reply_mail(None, ns, entry_id="ORIG", body="ok", reply_all=True)
    assert out["status"] == "sent"
    assert out["subject"] == "RE: Topic"
    assert reply.sent == 1 and reply.saved == 0
    assert original.calls == ["ReplyAll"]


def test_reply_to_plain_text_mail_stays_plain():
    reply, _, ns = _reply_setup(html=None, body_format=1)
    mail_mod.reply_mail(None, ns, entry_id="ORIG", body="Sure", save_only=True)
    assert reply.BodyFormat == 1
    assert reply.Body == "Sure\n\nFrom: Ann\nOriginal"


def test_plain_body_on_html_reply_does_not_flatten_quoted_history():
    reply, _, ns = _reply_setup()
    mail_mod.reply_mail(None, ns, entry_id="ORIG", body="a < b", save_only=True)
    assert reply.Body == ""  # never assigned: that would flatten the HTML
    assert "<p class=MsoNormal>a &lt; b</p>" in reply.HTMLBody


def test_reply_extra_recipients_subject_importance():
    reply, _, ns = _reply_setup()
    mail_mod.reply_mail(
        None, ns, entry_id="ORIG", body="x", save_only=True,
        to=["bob@x.com"], cc=["cat@x.com"], bcc=["dan@x.com"],
        subject="RE: New subject", importance="high",
    )
    types = {r.Address: r.Type for r in reply.Recipients}
    assert types == {"ann@x.com": 1, "bob@x.com": 1, "cat@x.com": 2, "dan@x.com": 3}
    assert reply.Subject == "RE: New subject"
    assert reply.Importance == 2


def test_bad_importance_rejected():
    _, _, ns = _reply_setup()
    with pytest.raises(OutlookError, match="importance"):
        mail_mod.reply_mail(None, ns, entry_id="ORIG", body="x", importance="urgent")


def test_draft_only_mode_forces_drafts(monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_DRAFT_ONLY", "1")
    reply, _, ns = _reply_setup()
    out = mail_mod.reply_mail(None, ns, entry_id="ORIG", body="x")
    assert out["status"] == "saved_to_drafts"
    assert reply.sent == 0
    assert any("Draft-only" in n for n in out["notes"])


def test_unresolved_recipient_blocks_send_but_not_draft():
    reply, _, ns = _reply_setup()
    reply.Recipients.unresolvable.add("Bobby")
    with pytest.raises(OutlookError, match="Bobby"):
        mail_mod.reply_mail(None, ns, entry_id="ORIG", body="x", to=["Bobby"])
    assert reply.sent == 0

    reply2, _, ns2 = _reply_setup()
    reply2.Recipients.unresolvable.add("Bobby")
    out = mail_mod.reply_mail(None, ns2, entry_id="ORIG", body="x", to=["Bobby"], save_only=True)
    assert out["status"] == "saved_to_drafts"
    assert any("Bobby" in n for n in out["notes"])


def test_forward_requires_recipients_unless_draft():
    fwd, _, ns = _reply_setup(recipients=())
    with pytest.raises(OutlookError, match="no recipients"):
        mail_mod.forward_mail(None, ns, entry_id="ORIG", body="FYI")
    out = mail_mod.forward_mail(None, ns, entry_id="ORIG", body="FYI", save_only=True)
    assert out["status"] == "saved_to_drafts"
    assert out["forwarded"] == "ORIG"


def test_forward_sends_with_note_on_top():
    fwd, original, ns = _reply_setup(recipients=())
    out = mail_mod.forward_mail(None, ns, entry_id="ORIG", to=["bob@x.com"], body="FYI")
    assert out["status"] == "sent"
    assert original.calls == ["Forward"]
    assert fwd.HTMLBody.index("FYI") < fwd.HTMLBody.index("From:")


def test_send_draft(monkeypatch):
    draft = FakeMail(html="<html><body>x</body></html>", entry_id="D1")
    draft.Recipients.Add("ann@x.com")
    ns = FakeNamespace({"D1": draft})
    out = mail_mod.send_draft(None, ns, entry_id="D1")
    assert out["status"] == "sent" and draft.sent == 1

    monkeypatch.setenv("OUTLOOK_MCP_DRAFT_ONLY", "yes")
    with pytest.raises(OutlookError, match="draft-only"):
        mail_mod.send_draft(None, ns, entry_id="D1")


def test_send_draft_refuses_sent_or_empty_items():
    sent = FakeMail(entry_id="S1")
    sent.Sent = True
    sent.Recipients.Add("a@x.com")
    empty = FakeMail(entry_id="E1")
    ns = FakeNamespace({"S1": sent, "E1": empty})
    with pytest.raises(OutlookError, match="already been sent"):
        mail_mod.send_draft(None, ns, entry_id="S1")
    with pytest.raises(OutlookError, match="no recipients"):
        mail_mod.send_draft(None, ns, entry_id="E1")


def _signature_dir(tmp_path, monkeypatch):
    d = tmp_path / "Signatures"
    (d / "Work_files").mkdir(parents=True)
    (d / "Work_files" / "image001.png").write_bytes(b"\x89PNG")
    (d / "Work.htm").write_bytes(
        b'<html><body><div class=WordSection1><p class=MsoNormal>Ori Harel</p>'
        b'<img src="Work_files/image001.png"></div></body></html>'
    )
    monkeypatch.setenv("OUTLOOK_MCP_SIGNATURES_DIR", str(d))


def test_reply_with_named_signature_embeds_image(tmp_path, monkeypatch):
    _signature_dir(tmp_path, monkeypatch)
    reply, _, ns = _reply_setup()
    out = mail_mod.reply_mail(None, ns, entry_id="ORIG", body="Thanks", signature="Work", save_only=True)
    html = reply.HTMLBody
    assert out["signature"] == "Work"
    assert html.index("Thanks") < html.index("Ori Harel") < html.index("From:")
    [att] = reply.Attachments.added
    cid = att.PropertyAccessor.props[mail_mod.PR_ATTACH_CONTENT_ID]
    assert f'src="cid:{cid}"' in html
    assert att.PropertyAccessor.props[mail_mod.PR_ATTACHMENT_HIDDEN] is True


def test_default_signature_not_duplicated_when_outlook_added_one():
    html = QUOTED.replace("<p class=MsoNormal>&nbsp;</p>", '<a name="_MailAutoSig">SIG</a>')
    reply, _, ns = _reply_setup(html=html)
    out = mail_mod.reply_mail(None, ns, entry_id="ORIG", body="x", signature=True, save_only=True)
    assert reply.HTMLBody.count("SIG") == 1
    assert any("already added" in n for n in out["notes"])


def test_default_reply_signature_missing_is_reported():
    reply, _, ns = _reply_setup()
    out = mail_mod.reply_mail(None, ns, entry_id="ORIG", body="x", signature="default", save_only=True)
    assert "signature" not in out
    assert any("No default" in n for n in out["notes"])


def test_new_mail_with_named_signature(tmp_path, monkeypatch):
    _signature_dir(tmp_path, monkeypatch)
    mail = FakeMail(entry_id="N1")
    out = mail_mod.send_mail(
        fake_outlook_creating(mail), FakeNamespace({}),
        to=["a@x.com"], subject="Hello", body="Line 1\n\nLine 2",
        signature="Work", save_only=True,
    )
    assert out["status"] == "saved_to_drafts"
    assert mail.BodyFormat == 2
    html = mail.HTMLBody
    assert html.index("Line 1") < html.index("Line 2") < html.index("Ori Harel")
    assert "<div class=WordSection1>" in html


def test_new_mail_plain_without_signature_unchanged():
    mail = FakeMail(entry_id="N2")
    mail_mod.send_mail(
        fake_outlook_creating(mail), FakeNamespace({}),
        to=["a@x.com"], subject="Hello", body="plain", save_only=True,
    )
    assert mail.BodyFormat == 1 and mail.Body == "plain"


def test_bad_attachment_rejected_before_item_is_created(tmp_path):
    created = []

    class Outlook:
        def CreateItem(self, kind):
            created.append(kind)
            return FakeMail()

    with pytest.raises(OutlookError):
        mail_mod.send_mail(
            Outlook(), FakeNamespace({}), to=["a@x.com"], subject="s", body="b",
            attachments=[str(tmp_path / "missing.pdf")],
        )
    assert created == []


def test_reply_to_rich_text_mail_converts_to_html_instead_of_flattening():
    reply, _, ns = _reply_setup(html="<html><body><p>Quoted RTF</p></body></html>", body_format=3)
    reply.Body = "Quoted RTF"
    mail_mod.reply_mail(None, ns, entry_id="ORIG", body="Answer", save_only=True)
    assert reply.BodyFormat == 2
    assert reply.Body == "Quoted RTF"  # .Body never overwritten
    assert reply.HTMLBody.index("Answer") < reply.HTMLBody.index("Quoted RTF")
