"""Quick Steps parsing/running, bulk move, and the Waiting check."""

import datetime as dt
from types import SimpleNamespace

import pytest

from outlook_mcp.client import quicksteps as qs
from outlook_mcp.client import triage
from outlook_mcp.errors import OutlookError

FID_ACTION = "00000000" + "AB" * 22
FID_WAIT = "00000000" + "CD" * 22

XML_ELEMENTS = f"""<?xml version="1.0"?>
<CombinedAction Ordinal="1" Name="Action">
  <ActionMoveToFolder><FolderEntryID>{FID_ACTION}</FolderEntryID></ActionMoveToFolder>
  <ActionMarkAsRead />
</CombinedAction>"""

XML_ATTRS = f"""<QuickStep Name="Waiting"><qakMoveToFolder Folder="{FID_WAIT}"/>
<qakCategorizeMessage Categories="{{11111111-2222-3333-4444-555555555555}}"/></QuickStep>"""

XML_REPLY = """<CombinedAction Name="Reply and Delete"><ActionReply/><ActionDeleteMessage/></CombinedAction>"""
XML_DELETE = """<CombinedAction Name="Done"><ActionMarkAsRead/><ActionDeleteMessage/></CombinedAction>"""
XML_UNKNOWN_FOLDER = """<CombinedAction Name="Read"><ActionMoveToFolder><FolderEntryID>zz</FolderEntryID></ActionMoveToFolder></CombinedAction>"""


def test_parse_element_style():
    out = qs.parse_quick_step_xml(XML_ELEMENTS)
    assert out["name"] == "Action"
    assert [a["kind"] for a in out["actions"]] == ["move", "mark_read"]
    assert out["actions"][0]["folder_ids"] == [FID_ACTION]


def test_parse_attribute_style_with_category_guid():
    out = qs.parse_quick_step_xml(XML_ATTRS)
    assert out["name"] == "Waiting"
    assert [a["kind"] for a in out["actions"]] == ["move", "categorize"]
    assert out["actions"][0]["folder_ids"] == [FID_WAIT]
    assert out["actions"][1]["category_refs"] == ["{11111111-2222-3333-4444-555555555555}"]


# ------------------------------------------------------------------ fakes


class _Folder:
    def __init__(self, name, path):
        self.Name, self.FolderPath, self.items = name, path, []


class _Msg:
    def __init__(self, eid, subject="s", **kw):
        self.EntryID, self.Subject, self.Class = eid, subject, 43
        self.UnRead, self.Categories, self.FlagStatus, self.Importance = True, "", 0, 1
        self.saved = self.deleted = 0
        self.Parent = None
        for k, v in kw.items():
            setattr(self, k, v)

    def Save(self):
        self.saved += 1

    def Move(self, folder):
        folder.items.append(self)
        self.Parent = folder
        return self

    def Delete(self):
        self.deleted += 1

    def Copy(self):
        return _Msg(self.EntryID + "-copy", self.Subject)


def _ns(xmls, msgs=()):
    folders = {FID_ACTION: _Folder("1-Action", "\\\\me\\Inbox\\1-Action"),
               FID_WAIT: _Folder("2-Waiting", "\\\\me\\Inbox\\2-Waiting")}
    cat = SimpleNamespace(Name="Waiting", CategoryID="{11111111-2222-3333-4444-555555555555}")
    by_id = {m.EntryID: m for m in msgs}

    def folder_from_id(fid, *a):
        if fid not in folders:
            raise Exception("not found")
        return folders[fid]

    ns = SimpleNamespace(
        DefaultStore=SimpleNamespace(StoreID="S"),
        GetFolderFromID=folder_from_id,
        GetItemFromID=lambda eid, *a: by_id[eid],
        Categories=SimpleNamespace(Count=1, Item=lambda i: cat),
    )
    return ns, folders


@pytest.fixture
def steps(monkeypatch):
    def install(*xmls):
        monkeypatch.setattr(qs, "_raw_quick_steps", lambda ns: [(f"subj{i}", x) for i, x in enumerate(xmls)])
    return install


def test_list_marks_runnable_and_problems(steps):
    steps(XML_ELEMENTS, XML_ATTRS, XML_REPLY, XML_UNKNOWN_FOLDER)
    ns, _ = _ns([])
    out = qs.list_quick_steps(None, ns)["quick_steps"]
    by = {s["name"]: s for s in out}
    assert by["Action"]["runnable"] and by["Action"]["actions"][0]["folder"].endswith("1-Action")
    assert by["Waiting"]["runnable"] and by["Waiting"]["actions"][1]["categories"] == ["Waiting"]
    assert not by["Reply and Delete"]["runnable"]
    assert any("reply" in p for p in by["Reply and Delete"]["problems"])
    assert not by["Read"]["runnable"] and "folder not understood" in by["Read"]["problems"][0]
    assert all("_actions" not in s for s in out)


def test_run_quick_step_moves_and_marks(steps):
    steps(XML_ELEMENTS)
    a, b = _Msg("A"), _Msg("B")
    ns, folders = _ns([], [a, b])
    out = qs.run_quick_step(None, ns, name="action", entry_ids=["A", "B"])
    assert out["status"] == "done" and out["count"] == 2
    assert folders[FID_ACTION].items == [a, b] and not a.UnRead and a.saved == 1


def test_run_refuses_unsafe_and_unknown(steps):
    steps(XML_REPLY, XML_UNKNOWN_FOLDER)
    ns, _ = _ns([], [_Msg("A")])
    with pytest.raises(OutlookError, match="aren't run"):
        qs.run_quick_step(None, ns, name="Reply and Delete", entry_ids=["A"])
    with pytest.raises(OutlookError, match="not understood"):
        qs.run_quick_step(None, ns, name="Read", entry_ids=["A"])
    with pytest.raises(OutlookError, match="Your Quick Steps"):
        qs.run_quick_step(None, ns, name="Nope", entry_ids=["A"])


def test_delete_step_needs_confirm(steps):
    steps(XML_DELETE)
    m = _Msg("A")
    ns, _ = _ns([], [m])
    pre = qs.run_quick_step(None, ns, name="Done", entry_ids=["A"])
    assert pre["status"] == "preview" and m.deleted == 0
    qs.run_quick_step(None, ns, name="Done", entry_ids=["A"], confirm=True)
    assert m.deleted == 1 and not m.UnRead


def test_move_mails_bulk(monkeypatch):
    target = _Folder("Archive", "\\\\me\\Archive")
    monkeypatch.setattr(triage, "resolve_folder", lambda ns, spec: target)
    a, b = _Msg("A"), _Msg("B")
    ns = SimpleNamespace(GetItemFromID=lambda eid, *x: {"A": a, "B": b}[eid])
    out = triage.move_mails(None, ns, entry_ids=["A", "B", "C"], target_folder="archive", mark_read=True)
    assert out["count"] == 2 and target.items == [a, b] and not a.UnRead
    assert out["errors"][0]["entry_id"] == "C"


# ----------------------------------------------------------------- waiting


def _t(day, hour):
    return dt.datetime(2026, 11, day, hour, tzinfo=dt.timezone.utc)


def _mail(eid, sender, when, body="", sent=True):
    m = _Msg(eid, "RE: Budget", SenderName=sender, SenderEmailAddress="", ReceivedTime=when,
             SentOn=when, Body=body, Sent=sent, To="Dana")
    return m


def test_check_waiting(monkeypatch):
    waiting1 = _mail("W1", "Ori Harel", _t(1, 9))
    waiting2 = _mail("W2", "Ori Harel", _t(2, 9))
    reply = _mail("R1", "Dana Levi", _t(3, 10), body="Approved!\n\nFrom: Ori\nSent: x\nold")
    my_followup = _mail("M1", "Ori Harel", _t(4, 10))
    draft = _mail("D1", "Dana Levi", _t(5, 10), sent=False)
    convs = {"W1": [waiting1, reply, my_followup, draft], "W2": [waiting2]}
    for m in (waiting1, waiting2):
        m.GetConversation = (lambda eid=m.EntryID: SimpleNamespace(
            GetRootItems=lambda: convs[eid][:1],
            GetChildren=lambda node, eid=eid: (lambda lst: lst[lst.index(node) + 1:lst.index(node) + 2])(convs[eid]),
        ))
    folder = _Folder("2-Waiting", "\\\\me\\Inbox\\2-Waiting")
    folder.Items = [waiting1, waiting2]
    me = SimpleNamespace(Name="Ori Harel", Address="ori@corp.com", AddressEntry=None)
    ns = SimpleNamespace(CurrentUser=me)
    monkeypatch.setattr(triage, "resolve_folder", lambda ns, spec: folder)

    out = triage.check_waiting(None, ns, folder="Inbox/2-Waiting")
    assert out["answered_count"] == 1 and out["still_waiting_count"] == 1
    ans = out["answered"][0]
    assert ans["entry_id"] == "W1" and [r["from"] for r in ans["replies"]] == ["Dana Levi"]
    assert ans["replies"][0]["preview"] == "Approved!"
    assert out["still_waiting"][0]["entry_id"] == "W2"
