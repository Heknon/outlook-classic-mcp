"""Quote stripping and whole-thread retrieval."""

from types import SimpleNamespace

from outlook_mcp.client import mail
from outlook_mcp.utils.quoting import split_quoted

OUTLOOK_EN = """Sounds good, let's do Thursday.

Ori

________________________________
From: Dana Levi <dana.levi@contoso.com>
Sent: Sunday, November 8, 2026 10:00
To: Ori Harel
Subject: RE: Q3 numbers

Earlier text
"""

OUTLOOK_HE = """מעולה, נדבר מחר.

________________________________
מאת: דנה לוי <dana.levi@contoso.com>
נשלח: יום ראשון 08 נובמבר 2026 10:00
אל: אורי הראל
נושא: RE: Q3

טקסט קודם
"""

OUTLOOK_NO_RULE = """Thanks!

From: Yossi Mizrahi
Sent: Monday
Subject: x

old"""

ORIGINAL = """Ok.
-----Original Message-----
From: a
old"""

GMAIL = """Works for me.

On Sun, Nov 8, 2026 at 10:00 AM Dana Levi <dana@x.com> wrote:
> old
> older"""


def test_split_outlook_english_cuts_at_rule():
    new, quoted = split_quoted(OUTLOOK_EN)
    assert quoted and new == "Sounds good, let's do Thursday.\n\nOri"


def test_split_outlook_hebrew():
    new, quoted = split_quoted(OUTLOOK_HE)
    assert quoted and new == "מעולה, נדבר מחר."


def test_split_other_formats():
    assert split_quoted(OUTLOOK_NO_RULE) == ("Thanks!", True)
    assert split_quoted(ORIGINAL) == ("Ok.", True)
    assert split_quoted(GMAIL) == ("Works for me.", True)
    assert split_quoted("> a\n> b\n") == ("", True)


def test_no_quote_and_lone_from_line_kept():
    assert split_quoted("Just text.\r\nMore.") == ("Just text.\nMore.", False)
    # "From:" alone (no Sent/Date line after it) is not a reply header.
    body = "Notes\nFrom: the field team\nnext steps follow"
    assert split_quoted(body) == (body, False)
    assert split_quoted("") == ("", False)


# --------------------------------------------------------------- get_thread


def _msg(eid, when, sender, body, folder="Inbox", sent=True):
    return SimpleNamespace(
        Class=43, EntryID=eid, SentOn=when, ReceivedTime=when, SenderName=sender,
        To="", CC="", Subject="RE: Q3", Body=body, UnRead=False, Sent=sent,
        Attachments=SimpleNamespace(Count=0), Parent=SimpleNamespace(Name=folder),
        ConversationTopic="Q3",
    )


class _Conv:
    def __init__(self, tree):
        self.tree = tree  # {parent_id or None: [children]}

    def GetRootItems(self):
        return self.tree[None]

    def GetChildren(self, node):
        return self.tree.get(node.EntryID, [])


def test_get_thread_walks_conversation_oldest_first():
    import datetime as dt

    t = lambda h: dt.datetime(2026, 11, 8, h, tzinfo=dt.timezone.utc)  # noqa: E731
    root = _msg("A", t(9), "Dana Levi", "Q3 numbers attached.")
    mine = _msg("B", t(10), "Ori Harel", OUTLOOK_EN, folder="Sent Items")
    hers = _msg("C", t(11), "Dana Levi", "Thursday works.\n\nFrom: Ori\nSent: x\nold")
    draft = _msg("D", t(12), "Ori Harel", "draft text", folder="Drafts", sent=False)
    conv = _Conv({None: [root], "A": [mine], "B": [hers, draft]})
    for m in (root, mine, hers, draft):
        m.GetConversation = lambda c=conv: c
    ns = SimpleNamespace(GetItemFromID=lambda eid, *a: {"A": root, "B": mine, "C": hers, "D": draft}[eid])

    out = mail.get_thread(None, ns, entry_id="C")
    assert out["source"] == "conversation" and out["total_in_thread"] == 4
    assert [m["entry_id"] for m in out["messages"]] == ["A", "B", "C", "D"]
    assert out["messages"][1]["body"] == "Sounds good, let's do Thursday.\n\nOri"
    assert out["messages"][1]["folder"] == "Sent Items"
    assert out["messages"][2]["quoted_history_removed"] is True
    assert out["messages"][3]["is_draft"] is True

    full = mail.get_thread(None, ns, entry_id="C", strip_quoted=False, limit=2)
    assert [m["entry_id"] for m in full["messages"]] == ["C", "D"]  # most recent kept
    assert "Earlier text" not in full["messages"][0]["body"]
    assert full["total_in_thread"] == 4


def test_get_thread_falls_back_to_topic_search():
    import datetime as dt

    a = _msg("A", dt.datetime(2026, 11, 8, 9), "Dana", "one")
    b = _msg("B", dt.datetime(2026, 11, 8, 10), "Ori", "two", folder="Sent Items")
    a.GetConversation = lambda: None
    filters = []

    class _Items(list):
        def Restrict(self, f):
            filters.append(f)
            return self

    folders = {6: SimpleNamespace(Name="Inbox", Items=_Items([a])), 5: SimpleNamespace(Name="Sent Items", Items=_Items([b]))}
    ns = SimpleNamespace(GetItemFromID=lambda eid, *x: a, GetDefaultFolder=lambda k: folders[k])
    out = mail.get_thread(None, ns, entry_id="A")
    assert out["source"] == "topic_search"
    assert [m["entry_id"] for m in out["messages"]] == ["A", "B"]
    assert "0x0070001F\" = 'Q3'" in filters[0]
