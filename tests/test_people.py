"""Ctrl+K-style name lookup and searching mail by person."""

from types import SimpleNamespace

from outlook_mcp.client import contacts, mail


class _ExUser:
    def __init__(self, smtp, title="Engineer", dept="R&D"):
        self.PrimarySmtpAddress = smtp
        self.JobTitle = title
        self.Department = dept
        self.OfficeLocation = "TLV"
        self.BusinessTelephoneNumber = "123"
        self.Alias = smtp.split("@")[0]
        self.CompanyName = "Corp"
        self.MobileTelephoneNumber = None


class _Entry:
    def __init__(self, name, smtp):
        self.Name = name
        self.Address = f"/o=Corp/ou=Exchange/cn=Recipients/cn={smtp.split('@')[0]}"
        self._exu = _ExUser(smtp)

    def GetExchangeUser(self):
        return self._exu


GAL = [
    _Entry("Dana Levi", "dana.levi@corp.com"),
    _Entry("Dana Cohen", "dcohen@corp.com"),
    _Entry("Yossi Mizrahi", "yossi@corp.com"),
]


class _Recipient:
    def __init__(self, name):
        self.name = name
        self.Resolved = False
        self.AddressEntry = None

    def Resolve(self):
        # Like Outlook: resolves only a unique match.
        hits = [e for e in GAL if self.name.lower() in (e.Name.lower(), e._exu.PrimarySmtpAddress)]
        if len(hits) == 1:
            self.Resolved = True
            self.AddressEntry = hits[0]
        return self.Resolved


class _Folder:
    def __init__(self, name, items=()):
        self.Name = name
        self.items = list(items)
        self.filters = []

    @property
    def Items(self):
        folder = self

        class _Items(list):
            def Restrict(self, f):
                folder.filters.append(f)
                return _Items(folder.items)

            def Sort(self, *a):
                pass

        return _Items(self.items)


class _NS:
    def __init__(self, folders=None):
        self.Stores = []
        self.folders = folders or {}

    def CreateRecipient(self, name):
        return _Recipient(name)

    def GetGlobalAddressList(self):
        return SimpleNamespace(AddressEntries=GAL)

    def GetDefaultFolder(self, kind):
        return self.folders[kind]


def test_unique_name_resolves_with_details():
    out = contacts.resolve_name(None, _NS(), name="Yossi Mizrahi")
    assert out["status"] == "resolved" and out["resolved"] is True
    assert out["smtp_address"] == "yossi@corp.com"
    assert out["person"]["department"] == "R&D"


def test_partial_name_lists_candidates_like_ctrl_k():
    out = contacts.resolve_name(None, _NS(), name="dana")
    assert out["status"] == "ambiguous"
    assert {c["smtp_address"] for c in out["candidates"]} == {"dana.levi@corp.com", "dcohen@corp.com"}


def test_partial_name_with_single_directory_hit_resolves():
    out = contacts.resolve_name(None, _NS(), name="mizra")
    assert out["status"] == "resolved" and out["smtp_address"] == "yossi@corp.com"


def test_alias_matches_via_exchange_address():
    out = contacts.resolve_name(None, _NS(), name="dcohen")
    assert out["status"] == "resolved" and out["smtp_address"] == "dcohen@corp.com"


def test_not_found():
    out = contacts.resolve_name(None, _NS(), name="nobody here")
    assert out["status"] == "not_found" and out["resolved"] is False


def _mail(subject, received):
    return SimpleNamespace(
        Class=43, EntryID=subject, Subject=subject, SenderName="Yossi Mizrahi",
        SenderEmailAddress="x", To="", ReceivedTime=None, UnRead=False, FlagStatus=0,
        Attachments=None, Importance=1, Body="", _received=received,
    )


def test_search_by_person_builds_filter_and_merges_folders():
    inbox = _Folder("Inbox", [_mail("from yossi", "2")])
    sent = _Folder("Sent Items", [_mail("to yossi", "1")])
    ns = _NS({6: inbox, 5: sent})
    out = mail.search_mails_by_person(None, ns, person="Yossi Mizrahi", since="2026-11-01T00:00:00")
    assert out["count"] == 2 and out["folders"] == ["Inbox", "Sent Items"]
    f = inbox.filters[0]
    assert f.startswith("@SQL=(")
    assert "'yossi@corp.com'" in f and "fromname\" = 'Yossi Mizrahi'" in f
    assert "displayto\" LIKE '%Yossi Mizrahi%'" in f
    assert "datereceived" in f


def test_search_by_person_from_only_skips_sent_and_to_clauses():
    inbox = _Folder("Inbox", [])
    ns = _NS({6: inbox})
    out = mail.search_mails_by_person(None, ns, person="yossi@corp.com", direction="from")
    assert out["folders"] == ["Inbox"]
    assert "displayto" not in inbox.filters[0]


def test_search_by_ambiguous_person_returns_candidates_without_searching():
    inbox = _Folder("Inbox", [_mail("x", "1")])
    out = mail.search_mails_by_person(None, _NS({6: inbox}), person="dana")
    assert out["status"] == "ambiguous" and out["count"] == 0 and len(out["candidates"]) == 2
    assert inbox.filters == []
