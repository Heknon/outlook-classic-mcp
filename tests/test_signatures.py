"""Tests for utils.signatures — loading Outlook signature files."""

import pytest

from outlook_mcp.errors import OutlookError
from outlook_mcp.utils import signatures as sigs


def _make_sig_dir(tmp_path, monkeypatch):
    d = tmp_path / "Signatures"
    files = d / "Work_files"
    files.mkdir(parents=True)
    (files / "image001.png").write_bytes(b"\x89PNG")
    html = (
        '<html><head><meta http-equiv=Content-Type content="text/html; charset=windows-1252">'
        "</head><body lang=EN-US><div class=WordSection1>"
        "<p class=MsoNormal>Ori Harel – Caf\xe9</p>"
        '<p><img width=100 src="Work_files/image001.png" alt=logo></p>'
        '<p><img src="Work_files/image001.png"></p>'
        '<p><img src="https://example.com/x.png"></p>'
        '<p><img src="Work_files/missing.png"></p>'
        '<p><img src="../outside.png"></p>'
        "</div></body></html>"
    )
    (d / "Work.htm").write_bytes(html.encode("windows-1252"))
    (d / "Short.htm").write_bytes(b"<html><body><p>Short</p></body></html>")
    monkeypatch.setenv("OUTLOOK_MCP_SIGNATURES_DIR", str(d))
    return d


def test_list_and_load_signature(tmp_path, monkeypatch):
    _make_sig_dir(tmp_path, monkeypatch)
    assert sigs.list_signatures() == ["Short", "Work"]
    sig = sigs.load_signature("work")  # case-insensitive
    assert sig.name == "Work"
    assert "Caf\xe9" in sig.html and "–" in sig.html  # decoded with declared charset
    assert "<body" not in sig.html and "<head" not in sig.html


def test_images_become_cid_once_per_file(tmp_path, monkeypatch):
    _make_sig_dir(tmp_path, monkeypatch)
    sig = sigs.load_signature("Work")
    assert len(sig.images) == 1
    path, cid = sig.images[0]
    assert path.endswith("image001.png")
    assert sig.html.count(f'src="cid:{cid}"') == 2
    # External, missing and escaping references are left alone.
    assert 'src="https://example.com/x.png"' in sig.html
    assert 'src="Work_files/missing.png"' in sig.html
    assert 'src="../outside.png"' in sig.html


def test_url_encoded_src(tmp_path):
    d = tmp_path / "Signatures"
    (d / "My Sig_files").mkdir(parents=True)
    (d / "My Sig_files" / "image001.jpg").write_bytes(b"x")
    sig = sigs.prepare_signature_html(
        '<body><img src="My%20Sig_files/image001.jpg"></body>', d, name="My Sig"
    )
    assert len(sig.images) == 1
    assert "cid:" in sig.html


def test_unknown_signature_lists_available(tmp_path, monkeypatch):
    _make_sig_dir(tmp_path, monkeypatch)
    with pytest.raises(OutlookError, match="Short, Work"):
        sigs.load_signature("nope")


def test_missing_folder(tmp_path, monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_SIGNATURES_DIR", str(tmp_path / "absent"))
    assert sigs.list_signatures() == []
    with pytest.raises(OutlookError, match="signatures folder not found"):
        sigs.load_signature("Work")


def test_decode_reg_string():
    assert sigs.decode_reg_string("Work\x00") == "Work"
    assert sigs.decode_reg_string("Work".encode("utf-16-le") + b"\x00\x00") == "Work"
    assert sigs.decode_reg_string(None) == ""


def test_decode_htm_variants():
    assert sigs.decode_htm(b"\xef\xbb\xbfcaf\xc3\xa9") == "caf\xe9"
    assert sigs.decode_htm('<meta charset="utf-8">caf\xe9'.encode("utf-8")).endswith("caf\xe9")
    assert sigs.decode_htm("caf\xe9".encode("windows-1252")) == "caf\xe9"
    assert sigs.decode_htm(b'<meta charset="bogus-cs">x').endswith("x")


def test_default_signature_name_without_registry():
    assert sigs.default_signature_name("new", "Outlook", ["a@x.com"]) is None
