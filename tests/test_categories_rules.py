"""Categories (colours) and rule creation, against small fakes."""

import datetime as dt
from types import SimpleNamespace

import pytest

from outlook_mcp.client import categories as cats
from outlook_mcp.client import rules
from outlook_mcp.errors import OutlookError
from outlook_mcp.schemas import RuleActions, RuleConditions

# ---------------------------------------------------------------- fakes


class _Cat:
    def __init__(self, name, color):
        self.Name, self.Color, self.CategoryID = name, color, "{" + name + "}"


class _Cats:
    def __init__(self, *items):
        self.items = list(items)

    @property
    def Count(self):
        return len(self.items)

    def Item(self, i):
        return self.items[i - 1]

    def Add(self, name, color=0, *a):
        c = _Cat(name, color)
        self.items.append(c)
        return c

    def Remove(self, key):
        self.items = [c for c in self.items if key not in (c.CategoryID, c.Name)]


class _Item:
    def __init__(self, eid, categories=""):
        self.EntryID, self.Categories, self.Subject, self.saved = eid, categories, f"subj {eid}", 0

    def Save(self):
        self.saved += 1


def _ns(items=(), categories=None, **kw):
    by_id = {i.EntryID: i for i in items}
    ns = SimpleNamespace(
        Categories=categories or _Cats(_Cat("Work", 8)),
        GetItemFromID=lambda eid, *a: by_id[eid],
        **kw,
    )
    return ns


# ---------------------------------------------------------- categories


def test_color_names_and_aliases():
    assert cats.color_code("Red") == 1
    assert cats.color_code("dark blue") == 23 == cats.color_code("navy") == cats.color_code("dark_blue")
    assert cats.color_code("grey") == 13
    assert cats.color_code(5) == 5
    with pytest.raises(OutlookError, match="Choose one of"):
        cats.color_code("sparkly")


def test_list_create_update_delete():
    ns = _ns()
    assert cats.list_categories(None, ns)["items"] == [{"name": "Work", "color": "blue", "color_code": 8}]
    out = cats.create_category(None, ns, name="Customer A", color="green")
    assert out == {"status": "created", "name": "Customer A", "color": "green", "color_code": 5}
    with pytest.raises(OutlookError, match="already exists"):
        cats.create_category(None, ns, name="customer a")
    with pytest.raises(OutlookError, match="separators"):
        cats.create_category(None, ns, name="a, b")
    out = cats.update_category(None, ns, name="Customer A", color="dark red", new_name="Cust A")
    assert out["name"] == "Cust A" and out["color"] == "dark red" and "note" in out
    cats.delete_category(None, ns, name="cust a")
    assert [c.Name for c in ns.Categories.items] == ["Work"]


def test_set_category_modes_and_auto_create():
    a, b = _Item("A", "Work"), _Item("B", "Work, Old")
    ns = _ns([a, b])
    out = cats.set_category(None, ns, entry_ids=["A", "B"], categories=["Urgent"], mode="add", color="red")
    assert a.Categories == "Work, Urgent" and b.Categories == "Work, Old, Urgent"
    assert out["created_categories"] == ["Urgent"]
    assert next(c for c in ns.Categories.items if c.Name == "Urgent").Color == 1
    cats.set_category(None, ns, entry_ids=["A", "B"], categories="urgent, old", mode="remove")
    assert a.Categories == "Work" and b.Categories == "Work"
    out = cats.set_category(None, ns, entry_id="A", categories="")
    assert a.Categories == "" and out["categories"] == ""  # single-item shape kept
    with pytest.raises(OutlookError, match="mode"):
        cats.set_category(None, ns, entry_id="A", categories="x", mode="toggle")


def test_set_category_bulk_reports_errors_per_item():
    a = _Item("A")
    ns = _ns([a])
    out = cats.set_category(None, ns, entry_ids=["A", "missing"], categories="Work")
    assert out["count"] == 1 and out["errors"][0]["entry_id"] == "missing"


# --------------------------------------------------------------- rules


class _Cond:
    def __init__(self):
        self.Enabled = False
        self.Recipients = _Recips()
        self.Text = ()
        self.Address = ()
        self.Categories = ()
        self.Importance = 1
        self.Folder = None
        self.MarkInterval = 4


class _Recips(list):
    def Add(self, who):
        r = SimpleNamespace(Name=who, Address=who, Resolved="?" not in who)
        self.append(r)
        return r

    def ResolveAll(self):
        return all(r.Resolved for r in self)


class _Side:
    def __getattr__(self, name):
        obj = _Cond()
        setattr(self, name, obj)
        return obj


class _Rule:
    def __init__(self, name):
        self.Name, self.Enabled, self.RuleType, self.IsLocalRule = name, True, 0, False
        self.Conditions, self.Exceptions, self.Actions = _Side(), _Side(), _Side()
        self.executed = []

    def Execute(self, *args):
        self.executed.append(args)


class _Rules:
    def __init__(self, store):
        self.store = store
        self.items = list(store.saved)

    @property
    def Count(self):
        return len(self.items)

    def Item(self, i):
        return self.items[i - 1]

    def Create(self, name, kind):
        r = _Rule(name)
        self.items.append(r)
        return r

    def Remove(self, i):
        del self.items[i - 1]

    def Save(self, *a):
        self.store.saved = list(self.items)
        self.store.saves += 1


class _Store:
    def __init__(self):
        self.saved, self.saves = [], 0

    def GetRules(self):
        return _Rules(self)


def _mail(subject, sender, smtp, days_ago=1, to="Ori Harel", attachments=0):
    rt = dt.datetime.now() - dt.timedelta(days=days_ago)
    return SimpleNamespace(Class=43, Subject=subject, SenderName=sender, SenderEmailAddress=smtp,
                           To=to, CC="", Body="", Importance=1, Categories="", ReceivedTime=rt,
                           Attachments=SimpleNamespace(Count=attachments))


class _Items(list):
    def Sort(self, *a):
        self.sort(key=lambda m: m.ReceivedTime, reverse=True)


def _rules_ns(monkeypatch, mails=()):
    store = _Store()
    inbox = SimpleNamespace(Name="Inbox", Items=_Items(mails))
    vendors = SimpleNamespace(Name="Vendors", FolderPath="\\\\ori\\Inbox\\Vendors")
    folders = {"inbox": inbox, "inbox/vendors": vendors}
    monkeypatch.setattr(rules, "resolve_folder", lambda ns, spec: folders[(spec or "inbox").lower()])
    me = SimpleNamespace(Name="Ori Harel", Address="ori@corp.com",
                         AddressEntry=SimpleNamespace(GetExchangeUser=lambda: SimpleNamespace(PrimarySmtpAddress="ori@corp.com")))

    def lookup(ns, who, *a):
        if "@" in who:
            return {"status": "resolved", "person": {"smtp_address": who, "display_name": who}}
        return {"status": "not_found"}

    import outlook_mcp.client.contacts as contacts

    monkeypatch.setattr(contacts, "lookup_person", lookup)
    return SimpleNamespace(DefaultStore=store, CurrentUser=me, Categories=_Cats()), store


def test_create_rule_previews_then_creates(monkeypatch):
    ns, store = _rules_ns(monkeypatch, [
        _mail("Invoice 1", "Acme", "billing@acme.com"),
        _mail("Hello", "Dana", "dana@corp.com"),
        _mail("Invoice old", "Acme", "billing@acme.com", days_ago=90),
    ])
    c = RuleConditions(sender_address_words=["@acme.com"], subject_words=["invoice"])
    a = RuleActions(move_to_folder="Inbox/Vendors", categories=["Vendors"], category_color="orange")
    pre = rules.create_rule(None, ns, name="Acme invoices", conditions=c, actions=a)
    assert pre["status"] == "preview" and store.saves == 0
    assert pre["matches"]["would_match"] == 1  # the 90-day-old one is outside the 30-day preview
    assert pre["rule"]["actions"] == {"move_to_folder": "Inbox/Vendors", "categories": ["Vendors"]}

    out = rules.create_rule(None, ns, name="Acme invoices", conditions=c, actions=a, confirm=True)
    assert out["status"] == "created" and store.saves == 1
    rule = store.saved[0]
    assert rule.Conditions.SenderAddress.Enabled and rule.Conditions.SenderAddress.Address == ("@acme.com",)
    assert rule.Conditions.Subject.Text == ("invoice",)
    assert rule.Actions.MoveToFolder.Folder.Name == "Vendors"
    assert rule.Actions.AssignToCategory.Categories == ("Vendors",)
    assert next(c for c in ns.Categories.items if c.Name == "Vendors").Color == 2
    assert out["rule"]["actions"]["move_to_folder"].endswith("Vendors")


def test_duplicate_name_needs_replace(monkeypatch):
    ns, store = _rules_ns(monkeypatch)
    c, a = RuleConditions(has_attachment=True), RuleActions(desktop_alert=True)
    rules.create_rule(None, ns, name="R", conditions=c, actions=a, confirm=True)
    with pytest.raises(OutlookError, match="already exists"):
        rules.create_rule(None, ns, name="r", conditions=c, actions=a, confirm=True)
    out = rules.create_rule(None, ns, name="R", conditions=c, actions=RuleActions(stop_processing=True),
                            replace=True, confirm=True)
    assert out["status"] == "replaced" and len(store.saved) == 1
    assert store.saved[0].Actions.Stop.Enabled


def test_forward_rules_blocked_by_default_and_external(monkeypatch):
    ns, store = _rules_ns(monkeypatch)
    c = RuleConditions(subject_words=["x"])
    monkeypatch.delenv("OUTLOOK_MCP_ALLOW_SEND", raising=False)
    with pytest.raises(OutlookError, match="draft-only"):
        rules.create_rule(None, ns, name="fw", conditions=c, actions=RuleActions(forward_to=["dana@corp.com"]))
    monkeypatch.setenv("OUTLOOK_MCP_ALLOW_SEND", "1")
    with pytest.raises(OutlookError, match="own organization"):
        rules.create_rule(None, ns, name="fw", conditions=c, actions=RuleActions(redirect_to=["me@gmail.com"]))
    ok = rules.create_rule(None, ns, name="fw", conditions=c, actions=RuleActions(forward_to=["dana@corp.com"]))
    assert ok["status"] == "preview"
    monkeypatch.setenv("OUTLOOK_MCP_ALLOW_EXTERNAL_FORWARD", "1")
    assert rules.create_rule(None, ns, name="fw2", conditions=c,
                             actions=RuleActions(forward_to=["me@gmail.com"]))["status"] == "preview"
    assert store.saves == 0


def test_validation(monkeypatch):
    ns, _ = _rules_ns(monkeypatch)
    with pytest.raises(OutlookError, match="at least one action"):
        rules.create_rule(None, ns, name="x", conditions=RuleConditions(), actions=RuleActions())
    with pytest.raises(OutlookError, match="not both"):
        rules.create_rule(None, ns, name="x", conditions=RuleConditions(),
                          actions=RuleActions(move_to_folder="inbox", delete=True))
    pre = rules.create_rule(None, ns, name="x", conditions=RuleConditions(), actions=RuleActions(delete=True))
    assert any("EVERY" in n for n in pre["notes"])


def test_unresolved_recipient_not_saved(monkeypatch):
    ns, store = _rules_ns(monkeypatch)
    with pytest.raises(OutlookError, match="resolve"):
        rules.create_rule(None, ns, name="x", conditions=RuleConditions(from_people=["Dana?"]),
                          actions=RuleActions(desktop_alert=True), confirm=True)
    assert store.saves == 0


def test_exceptions_in_preview(monkeypatch):
    ns, _ = _rules_ns(monkeypatch, [
        _mail("Report", "Acme", "billing@acme.com", attachments=1),
        _mail("Report", "Acme", "boss@acme.com", attachments=1),
    ])
    pre = rules.create_rule(None, ns, name="x", conditions=RuleConditions(has_attachment=True),
                            exceptions=RuleConditions(sender_address_words=["boss@"]),
                            actions=RuleActions(categories=["Reports"]))
    assert pre["matches"]["would_match"] == 1


def test_list_details_delete_and_run(monkeypatch):
    ns, store = _rules_ns(monkeypatch, [_mail("Invoice", "Acme", "billing@acme.com")])
    rules.create_rule(None, ns, name="Acme", conditions=RuleConditions(subject_words=["invoice"]),
                      actions=RuleActions(mark_as_task="tomorrow"), confirm=True)
    d = rules.list_rules(None, ns, details=True)["items"][0]
    assert d["conditions"] == {"subject_words": ["invoice"]} and d["actions"] == {"mark_as_task": "tomorrow"}

    pre = rules.run_rule(None, ns, name="acme")
    assert pre["status"] == "preview" and pre["matches"]["would_match"] == 1 and not store.saved[0].executed
    out = rules.run_rule(None, ns, name="acme", which="unread", confirm=True)
    assert out["status"] == "ran"
    assert store.saved[0].executed[0][2:] == (False, 2)

    assert rules.delete_rule(None, ns, name="Acme")["status"] == "preview" and len(store.saved) == 1
    rules.delete_rule(None, ns, name="Acme", confirm=True)
    assert store.saved == []
    with pytest.raises(OutlookError, match="not found"):
        rules.delete_rule(None, ns, name="Acme")
