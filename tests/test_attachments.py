"""Saving attachments: safe names, default folder, inline/embedded/link handling."""

import datetime as dt
import os
from types import SimpleNamespace

import pytest

from outlook_mcp.client import mail
from outlook_mcp.errors import OutlookError
from outlook_mcp.utils.attachments import folder_slug, kind_of, linked_files, safe_filename


def test_safe_filename_neutralizes_paths_and_bad_chars():
    assert safe_filename("Report 10:30.pdf", "x") == "Report 10_30.pdf"
    assert safe_filename("..\\..\\Windows\\evil.exe", "x") == "_.._Windows_evil.exe"
    assert safe_filename("../../etc/passwd", "x") == "_.._etc_passwd"
    assert safe_filename("C:\\temp\\a.txt", "x") == "C__temp_a.txt"
    assert safe_filename("..", "fallback") == "fallback"
    assert safe_filename("", "fallback") == "fallback"
    assert safe_filename("CON.txt", "x") == "_CON.txt"
    assert safe_filename("הצעת מחיר.pdf", "x") == "הצעת מחיר.pdf"
    long = "a" * 300 + ".pdf"
    assert len(safe_filename(long, "x")) == 150 and safe_filename(long, "x").endswith(".pdf")
    for raw in ("a/b", "a\\b", "c:d", "x\x00y"):
        assert os.sep not in safe_filename(raw, "f") and "/" not in safe_filename(raw, "f")


def test_kinds_and_slug():
    assert kind_of("Deck.PPTX") == "powerpoint" and kind_of("a.docx") == "word"
    assert kind_of("x.pdf") == "pdf" and kind_of("noext") == "other"
    assert folder_slug("RE: Q3 / numbers?") == "RE_ Q3 _ numbers"


def test_linked_files():
    body = (
        "See https://contoso.sharepoint.com/:w:/r/sites/Fin/Doc.docx?d=1, and "
        "http://portal.corp.local/sites/hr/Shared%20Documents/x.pptx. Also \\\\fs01\\share\\Budget.xlsx "
        "and https://example.com/not-a-file"
    )
    assert linked_files(body) == [
        "https://contoso.sharepoint.com/:w:/r/sites/Fin/Doc.docx?d=1",
        "http://portal.corp.local/sites/hr/Shared%20Documents/x.pptx",
        "\\\\fs01\\share\\Budget.xlsx",
    ]


class _PA:
    def __init__(self, hidden=False, cid=None):
        self.hidden, self.cid = hidden, cid

    def GetProperty(self, tag):
        if tag.endswith("7FFE000B"):
            return self.hidden
        if tag.endswith("3712001F"):
            if self.cid is None:
                raise Exception("not found")
            return self.cid
        raise Exception("not found")


class _Att:
    def __init__(self, name, data=b"x", type_=1, hidden=False, cid=None, fail=False, path=None):
        self.FileName, self.DisplayName, self.Type = name, name, type_
        self.Size, self.data, self.fail = len(data), data, fail
        self.PropertyAccessor = _PA(hidden, cid)
        if path:
            self.PathName = path

    def SaveAsFile(self, target):
        if self.fail:
            raise Exception("blocked by policy")
        with open(target, "wb") as f:
            f.write(self.data)


def _item(atts, body="", subject="Q3 numbers"):
    return SimpleNamespace(
        Attachments=atts, Body=body, Subject=subject,
        ReceivedTime=dt.datetime(2026, 11, 5, 9, 0, tzinfo=dt.timezone.utc),
    )


def test_save_default_folder_skips_inline_and_reports_links(monkeypatch, tmp_path):
    monkeypatch.setenv("OUTLOOK_MCP_ATTACHMENTS_DIR", str(tmp_path / "att"))
    item = _item(
        [
            _Att("Q3.pdf", b"%PDF-1"),
            _Att("image001.png", b"png", hidden=True, cid="image001@x"),
            _Att("Fwd note", b"msg", type_=5),
            _Att("Deck.pptx", type_=7, path="https://contoso.sharepoint.com/:p:/r/Deck.pptx"),
            _Att("Locked.xlsx", fail=True),
        ],
        body="Details: https://contoso.sharepoint.com/:w:/r/sites/Fin/Plan.docx",
    )
    ns = SimpleNamespace(GetItemFromID=lambda eid, *a: item)
    out = mail.save_attachments(None, ns, entry_id="E")
    assert out["output_dir"] == str((tmp_path / "att" / "2026-11-05 Q3 numbers").resolve())
    names = [(f["filename"], f["kind"]) for f in out["files"]]
    assert names == [("Q3.pdf", "pdf"), ("Fwd note.msg", "email")]
    assert open(out["files"][0]["path"], "rb").read() == b"%PDF-1"
    reasons = {s["filename"]: s["reason"] for s in out["skipped"]}
    assert reasons["image001.png"] == "inline image"
    assert reasons["Deck.pptx"] == "link to a file stored elsewhere"
    assert reasons["Locked.xlsx"].startswith("could not save")
    assert out["linked_files"] == ["https://contoso.sharepoint.com/:w:/r/sites/Fin/Plan.docx"]

    again = mail.save_attachments(None, ns, entry_id="E", attachment_index=2)  # inline, but asked for
    assert again["files"][0]["filename"] == "image001.png"
    # Saving the same attachment again reuses the identical file ...
    assert mail.save_attachments(None, ns, entry_id="E", attachment_indexes=[1])["files"][0]["filename"] == "Q3.pdf"
    # ... but a different file with the same name gets a numbered name.
    item.Attachments[0].data = b"%PDF-2 changed"
    assert mail.save_attachments(None, ns, entry_id="E", attachment_indexes=[1])["files"][0]["filename"] == "Q3 (1).pdf"
    assert not [p for p in os.listdir(out["output_dir"]) if p.endswith(".part")]
    with pytest.raises(OutlookError, match="out of range"):
        mail.save_attachments(None, ns, entry_id="E", attachment_index=9)


def test_explicit_output_dir_is_sandboxed(monkeypatch, tmp_path):
    monkeypatch.setenv("USERPROFILE", str(tmp_path / "profile"))
    (tmp_path / "profile").mkdir()
    ns = SimpleNamespace(GetItemFromID=lambda eid, *a: _item([_Att("a.pdf")]))
    with pytest.raises(OutlookError, match="outside the user profile"):
        mail.save_attachments(None, ns, entry_id="E", output_dir=str(tmp_path / "elsewhere"))
    out = mail.save_attachments(None, ns, entry_id="E", output_dir=str(tmp_path / "profile" / "Docs"))
    assert out["files"][0]["path"].startswith(str((tmp_path / "profile" / "Docs").resolve()))


def test_get_mail_lists_attachment_details():
    item = _item([_Att("Q3.pdf"), _Att("image001.png", hidden=True)], body="see \\\\fs01\\share\\x.xlsx")
    item.EntryID = "E"
    full = mail._mail_full(item)
    assert [(a["filename"], a["kind"], a["inline"]) for a in full["attachments"]] == [
        ("Q3.pdf", "pdf", False), ("image001.png", "image", True)]
    assert full["file_attachment_count"] == 1
    assert full["linked_files"] == ["\\\\fs01\\share\\x.xlsx"]
