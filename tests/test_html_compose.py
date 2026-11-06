"""Tests for utils.html_compose — where new text goes in Outlook HTML."""

from outlook_mcp.utils.html_compose import (
    EMPTY_PARAGRAPH,
    append_to_body,
    body_inner_html,
    compose_fragment,
    has_auto_signature,
    insert_at_top,
    new_html_document,
    text_to_html,
)

REPLY_HTML = (
    "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\"><head>"
    "<style>p.MsoNormal{font-family:Aptos;font-size:12pt}</style></head>"
    "<body lang=EN-US link=\"#467886\"><div class=WordSection1>"
    "<p class=MsoNormal><o:p>&nbsp;</o:p></p>"
    "<div style='border:none;border-top:solid #E1E1E1 1.0pt'><p><b>From:</b> Ann</p></div>"
    "<p class=MsoNormal>Original text</p></div></body></html>"
)


def test_text_to_html_escapes_and_keeps_blank_lines():
    out = text_to_html("Hi <Bob> & co\n\nThanks")
    assert out == (
        "<p class=MsoNormal>Hi &lt;Bob&gt; &amp; co</p>"
        + EMPTY_PARAGRAPH
        + "<p class=MsoNormal>Thanks</p>"
    )


def test_text_to_html_handles_crlf_and_empty():
    assert text_to_html("") == ""
    assert text_to_html("a\r\nb").count("<p class=MsoNormal>") == 2


def test_insert_goes_inside_wordsection_before_quoted_history():
    out = insert_at_top(REPLY_HTML, "<p class=MsoNormal>NEW</p>")
    assert out.index("NEW") > out.index("<div class=WordSection1>")
    assert out.index("NEW") < out.index("From:")
    assert out.index("NEW") > out.index("<body")
    # Head styles and quoted history untouched.
    assert "p.MsoNormal{font-family:Aptos" in out
    assert out.replace("<p class=MsoNormal>NEW</p>", "") == REPLY_HTML


def test_insert_without_wordsection_goes_right_after_body():
    html = "<html><body class=x><p>old</p></body></html>"
    assert insert_at_top(html, "N") == "<html><body class=x>N<p>old</p></body></html>"


def test_insert_into_fragment_or_empty():
    assert insert_at_top("<p>old</p>", "N") == "N<p>old</p>"
    assert insert_at_top("", "N") == "<html><body>N</body></html>"
    assert insert_at_top(None, "N") == "<html><body>N</body></html>"


def test_wordsection_with_quoted_class_attribute():
    html = '<body><div class="WordSection1"><p>q</p></div></body>'
    out = insert_at_top(html, "N")
    assert out == '<body><div class="WordSection1">N<p>q</p></div></body>'


def test_append_to_body():
    assert append_to_body("<html><body><p>a</p></body></html>", "S") == (
        "<html><body><p>a</p>S</body></html>"
    )
    assert append_to_body("<p>a</p>", "S") == "<p>a</p>S"


def test_compose_fragment_orders_body_then_signature():
    out = compose_fragment("Hi", is_html=False, signature_html="<p>SIG</p>", trailing_gap=True)
    assert out.index("Hi") < out.index("SIG")
    assert out.endswith(EMPTY_PARAGRAPH)


def test_compose_fragment_html_passthrough_and_empty():
    assert compose_fragment("<b>x</b>", is_html=True) == "<b>x</b>"
    assert compose_fragment("", is_html=False, trailing_gap=True) == ""


def test_body_inner_html_and_auto_signature():
    assert body_inner_html("<html><body a=1>X<p>y</p></body></html>") == "X<p>y</p>"
    assert body_inner_html("no body") == "no body"
    assert has_auto_signature('<a name="_MailAutoSig">') is True
    assert has_auto_signature("") is False
    assert has_auto_signature(None) is False


def test_new_html_document_wraps_fragment():
    doc = new_html_document("<p>x</p>")
    assert "<div class=WordSection1><p>x</p></div>" in doc
    assert "charset=utf-8" in doc
