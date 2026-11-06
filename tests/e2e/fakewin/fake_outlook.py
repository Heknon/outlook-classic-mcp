"""A small simulated Outlook object model for end-to-end tests.

Loaded instead of pywin32 (see the sibling ``pythoncom`` and
``win32com`` stubs), so the real MCP server can run over stdio on any OS
against a seeded mailbox. It mimics the behaviours the server relies on:

* pywin32 date conversion (aware datetimes are converted to UTC and that
  clock is stored as local wall time; values come back as local wall
  clock tagged UTC) — so timezone bugs show up;
* Restrict() with the DASL / Jet filters the server builds, evaluated for
  real (DASL dates in UTC, LIKE with %, AND/OR, parentheses);
* Reply/ReplyAll/Forward producing Word-style HTML with quoted history;
  Save() into Drafts, Send() into Sent Items;
* the address book (Resolve, GAL, Exchange users), shared calendars and
  Recipient.FreeBusy().

It is a test double, not Outlook: it only covers what the tools use.
"""

from __future__ import annotations

import datetime as dt
import itertools
import json
import os
import re

UTC = dt.timezone.utc
_ids = itertools.count(1)


def _new_id(prefix="ID"):
    return f"{prefix}{next(_ids):06d}"


# ----------------------------------------------------------- date handling


def _store(value):
    """Emulate pywin32 assigning a Python datetime to a COM DATE."""
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        if value.tzinfo is None:
            value = value.astimezone()  # naive is taken as local
        return value.astimezone(UTC).replace(tzinfo=None)  # UTC clock stored as wall time
    return value


def _load(value):
    """Emulate pywin32 returning a COM DATE: local wall clock tagged UTC."""
    if value is None:
        return None
    return value.replace(tzinfo=UTC)


def _utc(naive_local):
    return naive_local.astimezone(UTC).replace(tzinfo=None)


# ------------------------------------------------------------ directory


class ExchangeUser:
    def __init__(self, name, smtp, title="", dept="", office="", phone=""):
        self.Name = name
        self.PrimarySmtpAddress = smtp
        self.JobTitle = title
        self.Department = dept
        self.OfficeLocation = office
        self.BusinessTelephoneNumber = phone
        self.MobileTelephoneNumber = None
        self.CompanyName = "Contoso"
        self.Alias = smtp.split("@")[0]


class AddressEntry:
    def __init__(self, user: ExchangeUser):
        self._u = user
        self.Name = user.Name
        self.Address = f"/o=Contoso/ou=Exchange/cn=Recipients/cn={user.Alias}"

    def GetExchangeUser(self):
        return self._u


class Recipient:
    def __init__(self, app, text, type_=1):
        self._app = app
        self.Name = text
        self.Address = text
        self.Type = type_
        self.Resolved = False
        self.AddressEntry = None
        self.MeetingResponseStatus = 0

    def Resolve(self):
        hits = self._app._directory_lookup(self.Name)
        if len(hits) == 1:
            entry = hits[0]
            self.Resolved = True
            self.AddressEntry = entry
            self.Name = entry.Name
            self.Address = entry.Address
        elif "@" in self.Name and not hits:
            # External SMTP addresses resolve as one-off entries.
            self.Resolved = True
            self.Address = self.Name
            self.AddressEntry = type("OneOff", (), {"Name": self.Name, "Address": self.Name,
                                                    "GetExchangeUser": lambda s: None})()
        return self.Resolved

    def FreeBusy(self, start, interval, complete=True):
        if not self.Resolved:
            self.Resolve()
        user = getattr(self.AddressEntry, "_u", None)
        if user is None:
            raise Exception("The free/busy information could not be retrieved.")
        origin = _store(start)
        origin = origin.replace(hour=0, minute=0, second=0, microsecond=0)
        slots = 30 * 24 * 60 // interval
        codes = ["0"] * slots
        for appt in self._app._calendar_of(user.PrimarySmtpAddress):
            s, e = appt._start, appt._end
            for i in range(slots):
                t0 = origin + dt.timedelta(minutes=i * interval)
                t1 = t0 + dt.timedelta(minutes=interval)
                if s < t1 and e > t0:
                    code = str(appt.BusyStatus)
                    if int(code) > int(codes[i]) or codes[i] == "0":
                        codes[i] = code
        return "".join(codes)


class Recipients:
    def __init__(self, app):
        self._app = app
        self._items: list[Recipient] = []

    def Add(self, text):
        r = Recipient(self._app, text)
        self._items.append(r)
        return r

    def ResolveAll(self):
        return all([r.Resolve() or r.Resolved for r in self._items])

    def Item(self, i):
        return self._items[i - 1]

    def Remove(self, i):
        del self._items[i - 1]

    @property
    def Count(self):
        return len(self._items)

    def __iter__(self):
        return iter(list(self._items))

    def _display(self, kind):
        return "; ".join(r.Name for r in self._items if r.Type == kind)


class PropertyAccessor:
    def __init__(self):
        self.props = {}

    def SetProperty(self, name, value):
        self.props[name] = value


class Attachment:
    def __init__(self, path, display=None):
        self.FileName = os.path.basename(path)
        self.DisplayName = display or self.FileName
        self.Size = os.path.getsize(path) if os.path.exists(path) else 0
        self._path = path
        self.PropertyAccessor = PropertyAccessor()

    def SaveAsFile(self, target):
        with open(self._path, "rb") as src, open(target, "wb") as dst:
            dst.write(src.read())


class Attachments:
    def __init__(self):
        self._items: list[Attachment] = []

    def Add(self, path, *args):
        a = Attachment(path)
        self._items.append(a)
        return a

    @property
    def Count(self):
        return len(self._items)

    def __iter__(self):
        return iter(list(self._items))


# ---------------------------------------------------------------- items


class _Item:
    Class = 0

    def __init__(self, app):
        self._app = app
        self.EntryID = _new_id()
        self.Parent = None
        self.Categories = ""
        self.Subject = ""
        self.Body = ""
        self.Importance = 1

    def Delete(self):
        if self.Parent is not None:
            self.Parent._remove(self)
        self._app._deleted.append(self)

    def Move(self, folder):
        if self.Parent is not None:
            self.Parent._remove(self)
        folder._add(self)
        return self


def _quote_html(m: "MailItem") -> str:
    style = "<style>p.MsoNormal{margin:0cm;font-family:Aptos;font-size:12.0pt}</style>"
    header = (
        "<div style='border:none;border-top:solid #E1E1E1 1.0pt'><p class=MsoNormal>"
        f"<b>From:</b> {m.SenderName} <br><b>Subject:</b> {m.Subject}</p></div>"
    )
    original = m.HTMLBody or f"<p>{m.Body}</p>"
    inner = re.sub(r"(?is).*<body[^>]*>|</body>.*", "", original)
    return (
        f"<html><head>{style}</head><body lang=EN-US><div class=WordSection1>"
        f"<p class=MsoNormal>&nbsp;</p>{header}{inner}</div></body></html>"
    )


class MailItem(_Item):
    Class = 43

    def __init__(self, app, **kw):
        super().__init__(app)
        self.BodyFormat = 1
        self.HTMLBody = ""
        self.SenderName = ""
        self.SenderEmailAddress = ""
        self._sender_smtp = ""
        self.Recipients = Recipients(app)
        self._display = {1: "", 2: "", 3: ""}
        self.UnRead = False
        self.FlagStatus = 0
        self.Sent = False
        self.ConversationID = _new_id("CONV")
        self._received = None
        self._sent_on = None
        self.Attachments = Attachments()
        self.SendUsingAccount = None
        self.PropertyAccessor = PropertyAccessor()
        for k, v in kw.items():
            setattr(self, k, v)

    @property
    def ReceivedTime(self):
        return _load(self._received)

    # Like Outlook, switching a plain-text item to HTML converts its text.
    @property
    def BodyFormat(self):
        return self._fmt

    @BodyFormat.setter
    def BodyFormat(self, value):
        old = getattr(self, "_fmt", None)
        self._fmt = value
        if old == 1 and value == 2 and self.Body and not self.HTMLBody:
            paras = "".join(f"<p class=MsoNormal>{line or '&nbsp;'}</p>" for line in self.Body.split("\n"))
            self.HTMLBody = f"<html><body><div class=WordSection1>{paras}</div></body></html>"

    # Like Outlook: assigning To/CC/BCC replaces those recipients; reading
    # them gives the display string. Received mail keeps its seeded strings.
    def _set_line(self, kind, value):
        self.Recipients._items = [r for r in self.Recipients._items if r.Type != kind]
        for part in str(value or "").split(";"):
            if part.strip():
                self.Recipients.Add(part.strip()).Type = kind
        self._display[kind] = value or ""

    def _get_line(self, kind):
        if any(r.Type == kind for r in self.Recipients):
            return self.Recipients._display(kind)
        return self._display[kind]

    To = property(lambda s: s._get_line(1), lambda s, v: s._set_line(1, v))
    CC = property(lambda s: s._get_line(2), lambda s, v: s._set_line(2, v))
    BCC = property(lambda s: s._get_line(3), lambda s, v: s._set_line(3, v))

    @property
    def SentOn(self):
        return _load(self._sent_on)

    def _new_response(self, prefix, to_sender, all_=False):
        r = MailItem(self._app)
        r.Subject = f"{prefix}: {self.Subject}"
        r.ConversationID = self.ConversationID
        r.BodyFormat = self.BodyFormat if self.BodyFormat != 1 else 1
        if self.BodyFormat == 1:
            r.Body = f"\n\nFrom: {self.SenderName}\nSubject: {self.Subject}\n\n{self.Body}"
        else:
            r.HTMLBody = _quote_html(self)
            r.Body = re.sub("<[^>]+>", "", r.HTMLBody)
        if to_sender:
            r.Recipients.Add(self._sender_smtp or self.SenderEmailAddress)
        if all_:
            me = self._app._me.PrimarySmtpAddress.lower()
            for part in (self.To + ";" + self.CC).split(";"):
                part = part.strip()
                if part and part.lower() != me:
                    rec = r.Recipients.Add(part)
                    if part in self.CC:
                        rec.Type = 2
        if self.Attachments.Count and prefix == "FW":
            for a in self.Attachments:
                r.Attachments._items.append(a)
        return r

    def GetConversation(self):
        app = self._app
        conv_id = self.ConversationID

        def all_folders():
            stack = [app._root] + [st.GetRootFolder() for st in app._extra_stores]
            while stack:
                f = stack.pop()
                yield f
                stack.extend(f._subs)

        members = [i for f in all_folders() for i in f._items
                   if getattr(i, "ConversationID", None) == conv_id]
        members.sort(key=lambda i: i._received or i._sent_on or dt.datetime.max)

        class _Conversation:  # linear thread: each message answers the previous
            def GetRootItems(self):
                return members[:1]

            def GetChildren(self, node):
                ids = [m.EntryID for m in members]
                i = ids.index(node.EntryID) if node.EntryID in ids else -1
                return members[i + 1:i + 2] if i >= 0 else []

        return _Conversation()

    @property
    def ConversationTopic(self):
        return re.sub(r"^(?:(?:RE|FW|AW|WG)\s*:\s*)+", "", self.Subject, flags=re.I)

    def Reply(self):
        return self._new_response("RE", True)

    def ReplyAll(self):
        return self._new_response("RE", True, all_=True)

    def Forward(self):
        return self._new_response("FW", False)

    def _sync_display(self):
        pass

    def Save(self):
        self._sync_display()
        if self.Parent is None:
            self._app._folders[16]._add(self)  # Drafts

    def Send(self):
        self._sync_display()
        if not self.Recipients.ResolveAll():
            raise Exception("Outlook does not recognize one or more names.")
        if self.Parent is not None:
            self.Parent._remove(self)
        self.Sent = True
        self._sent_on = dt.datetime.now().replace(microsecond=0)
        self._received = self._sent_on
        self.SenderName = self._app._me.Name
        self._sender_smtp = self._app._me.PrimarySmtpAddress
        self._app._folders[5]._add(self)  # Sent Items
        self._app._outbox.append({"kind": "mail", "subject": self.Subject, "to": self.To})


class MeetingResponse(MailItem):
    pass


class AppointmentItem(_Item):
    Class = 26

    def __init__(self, app, **kw):
        super().__init__(app)
        self._start = None
        self._end = None
        self.AllDayEvent = False
        self.Location = ""
        self.Organizer = app._me.Name
        self.MeetingStatus = 0
        self.ResponseStatus = 1  # organized
        self.RecurrenceState = 0
        self.IsRecurring = False
        self.ReminderSet = False
        self.ReminderMinutesBeforeStart = 15
        self.BusyStatus = 2
        self.Recipients = Recipients(app)
        self._owner = app._me.PrimarySmtpAddress
        for k, v in kw.items():
            setattr(self, k, v)

    Start = property(lambda s: _load(s._start), lambda s, v: setattr(s, "_start", _store(v)))
    End = property(lambda s: _load(s._end), lambda s, v: setattr(s, "_end", _store(v)))

    def Save(self):
        if self.Parent is None:
            self._app._folders[9]._add(self)

    def Send(self):
        self._app._outbox.append(
            {"kind": "meeting", "subject": self.Subject, "status": self.MeetingStatus,
             "to": [r.Name for r in self.Recipients]}
        )

    def Respond(self, code, no_ui=True, *a):
        self.ResponseStatus = code
        resp = MeetingResponse(self._app)
        resp.Subject = {3: "Accepted", 2: "Tentative", 4: "Declined"}[code] + ": " + self.Subject
        resp.Recipients.Add(self.Organizer)
        return resp

    def GetRecurrencePattern(self):
        appt = self

        class _Pattern:
            def GetOccurrence(self, when):
                when = _store(when)
                for occ in appt._app._occurrences.get(appt.EntryID, []):
                    if occ._start == when:
                        return occ
                raise Exception("You changed one of the recurrences of this item, ...")

        return _Pattern()


class TaskItem(_Item):
    Class = 48

    def __init__(self, app):
        super().__init__(app)
        self._due = None
        self.Complete = False
        self.Status = 0
        self.PercentComplete = 0
        self.ReminderSet = False
        self._reminder = None

    DueDate = property(lambda s: _load(s._due), lambda s, v: setattr(s, "_due", _store(v)))
    ReminderTime = property(lambda s: _load(s._reminder), lambda s, v: setattr(s, "_reminder", _store(v)))

    def Save(self):
        if self.Parent is None:
            self._app._folders[13]._add(self)

    def MarkComplete(self):
        self.Complete = True


# ------------------------------------------------------------- filters


_PROP = {
    "urn:schemas:httpmail:subject": "Subject",
    "urn:schemas:httpmail:textdescription": "Body",
    "urn:schemas:httpmail:fromemail": "SenderEmailAddress",
    "urn:schemas:httpmail:fromname": "SenderName",
    "urn:schemas:httpmail:displayto": "To",
    "urn:schemas:httpmail:displaycc": "CC",
    "urn:schemas:httpmail:datereceived": "@received",
    "urn:schemas:calendar:dtstart": "@start",
    "http://schemas.microsoft.com/mapi/proptag/0x5D01001F": "_sender_smtp",
    "http://schemas.microsoft.com/mapi/proptag/0x5D02001F": "_sender_smtp",
}

_TOKEN = re.compile(
    r"""\s*(?:(?P<lp>\()|(?P<rp>\))|(?P<name>"[^"]+"|\[[^\]]+\])|(?P<str>'(?:[^']|'')*')"""
    r"""|(?P<op>>=|<=|<>|=|<|>|\bLIKE\b)|(?P<kw>\bAND\b|\bOR\b|\bTrue\b|\bFalse\b))""",
    re.IGNORECASE,
)


def _tokens(expr):
    pos, out = 0, []
    while pos < len(expr):
        if expr[pos:].strip() == "":
            break
        m = _TOKEN.match(expr, pos)
        if not m:
            raise Exception(f"Cannot parse filter near: {expr[pos:pos + 40]!r}")
        pos = m.end()
        kind = m.lastgroup
        out.append((kind, m.group(kind)))
    return out


class _Parser:
    def __init__(self, toks, dasl):
        self.t, self.i, self.dasl = toks, 0, dasl

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self, kind=None):
        tok = self.peek()
        if kind and tok[0] != kind:
            raise Exception(f"Filter syntax error: expected {kind}, got {tok}")
        self.i += 1
        return tok

    def parse(self):
        node = self.expr()
        if self.i != len(self.t):
            raise Exception(f"Filter syntax error at {self.t[self.i:]}")
        return node

    def expr(self):
        left = self.term()
        while self.peek()[0] == "kw" and self.peek()[1].upper() == "OR":
            self.take()
            right = self.term()
            left = ("or", left, right)
        return left

    def term(self):
        left = self.factor()
        while self.peek()[0] == "kw" and self.peek()[1].upper() == "AND":
            self.take()
            right = self.factor()
            left = ("and", left, right)
        return left

    def factor(self):
        if self.peek()[0] == "lp":
            self.take()
            node = self.expr()
            self.take("rp")
            return node
        _, name = self.take("name")
        _, op = self.take("op")
        kind, val = self.take()
        if kind == "str":
            val = val[1:-1].replace("''", "'")
        elif kind == "kw":
            val = val.lower() == "true"
        else:
            raise Exception(f"Filter syntax error: bad value {val}")
        return ("cmp", name, op.upper(), val)


def _value(item, name, dasl):
    if dasl:
        if not name.startswith('"'):
            raise Exception("DASL filters need quoted property names.")
        key = _PROP.get(name.strip('"'))
        if key is None:
            raise Exception(f"Unknown DASL property {name}")
        if key == "@received":
            return _utc(item._received) if item._received else None
        if key == "@start":
            return _utc(item._start) if item._start else None
        return getattr(item, key, None)
    if not name.startswith("["):
        raise Exception("Jet filters need [bracketed] property names.")
    prop = name[1:-1]
    return getattr(item, prop, None)


def _eval(node, item, dasl):
    if node[0] == "and":
        return _eval(node[1], item, dasl) and _eval(node[2], item, dasl)
    if node[0] == "or":
        return _eval(node[1], item, dasl) or _eval(node[2], item, dasl)
    _, name, op, val = node
    cur = _value(item, name, dasl)
    if isinstance(cur, dt.datetime) or (isinstance(val, str) and re.fullmatch(r"\d{4}-\d\d-\d\d \d\d:\d\d", val)):
        if cur is None:
            return False
        if not dasl:
            raise Exception("Jet date literals are locale-dependent; this test double rejects them.")
        val = dt.datetime.strptime(val, "%Y-%m-%d %H:%M")
    if op == "LIKE":
        pattern = "^" + ".*".join(re.escape(p) for p in str(val).lower().split("%")) + "$"
        return re.match(pattern, str(cur or "").lower()) is not None
    if isinstance(cur, str):
        cur, val = cur.lower(), str(val).lower()
    return {
        "=": cur == val, "<>": cur != val, ">=": cur >= val, "<=": cur <= val,
        ">": cur > val, "<": cur < val,
    }[op]


def compile_filter(f):
    dasl = f.startswith("@SQL=")
    expr = f[5:] if dasl else f
    tree = _Parser(_tokens(expr), dasl).parse()
    return lambda item: _eval(tree, item, dasl)


# ------------------------------------------------------------- folders


class Items:
    def __init__(self, folder, items=None):
        self._folder = folder
        self._items = list(folder._items if items is None else items)
        self.IncludeRecurrences = False
        self._sorted = None

    def Sort(self, prop, descending=False):
        key = prop.strip("[]")
        attr = {"ReceivedTime": "_received", "Start": "_start"}.get(key, key)
        self._items.sort(key=lambda i: getattr(i, attr, None) or dt.datetime.min, reverse=descending)

    def _expanded(self):
        if not self.IncludeRecurrences:
            return self._items
        out = []
        for it in self._items:
            occ = self._folder._app._occurrences.get(it.EntryID)
            out.extend(occ if occ else [it])
        out.sort(key=lambda i: i._start or dt.datetime.min)
        return out

    def Restrict(self, f):
        pred = compile_filter(f)
        res = Items(self._folder, [i for i in self._expanded() if pred(i)])
        self._folder._app._filters.append(f)
        return res

    def Add(self, kind=0):
        item = self._folder._app.CreateItem(kind)
        item.Parent = self._folder
        self._folder._items.append(item)
        return item

    @property
    def Count(self):
        return len(self._items)

    def __iter__(self):
        return iter(self._expanded())


class Folder:
    def __init__(self, app, name, kind, default_type=0, store_id="STORE-ME"):
        self._app = app
        self.Name = name
        self._kind = kind
        self.DefaultItemType = default_type
        self.EntryID = _new_id("FOLDER")
        self.StoreID = store_id
        self._items: list = []
        self._subs: list[Folder] = []
        self.Parent = None

    @property
    def Items(self):
        return Items(self)

    @property
    def Folders(self):
        folder = self

        class _Folders(list):
            def Add(self, name):
                f = Folder(folder._app, name, None)
                f.Parent = folder
                folder._subs.append(f)
                return f

        return _Folders(self._subs)

    @property
    def UnReadItemCount(self):
        return sum(1 for i in self._items if getattr(i, "UnRead", False))

    def _add(self, item):
        item.Parent = self
        self._items.append(item)
        self._app._by_id[item.EntryID] = item

    def _remove(self, item):
        if item in self._items:
            self._items.remove(item)


class Account:
    def __init__(self, app):
        self.DisplayName = app._me.PrimarySmtpAddress
        self.SmtpAddress = app._me.PrimarySmtpAddress
        self.UserName = app._me.Name
        self.AccountType = 0
        self.DeliveryStore = app._store


class _Coll(list):
    @property
    def Count(self):
        return len(self)

    def Item(self, i):
        return self[i - 1]


class _OleObj:
    def GetIDsOfNames(self, name):
        return 1


class Namespace:
    def __init__(self, app):
        self._app = app
        self._oleobj_ = _OleObj()
        self.CurrentProfileName = "Outlook"
        self.Offline = False
        self.ExchangeConnectionMode = 700

    @property
    def CurrentUser(self):
        r = Recipient(self._app, self._app._me.PrimarySmtpAddress)
        r.Resolve()
        return r

    @property
    def Accounts(self):
        return _Coll([Account(self._app)])

    @property
    def DefaultStore(self):
        return self._app._store

    @property
    def Stores(self):
        return _Coll([self._app._store, *self._app._extra_stores])

    def GetDefaultFolder(self, kind):
        return self._app._folders[kind]

    def GetItemFromID(self, entry_id, store_id=None):
        try:
            return self._app._by_id[entry_id]
        except KeyError:
            raise Exception("The operation failed. An object could not be found.") from None

    def CreateRecipient(self, text):
        return Recipient(self._app, text)

    def GetSharedDefaultFolder(self, recip, kind):
        smtp = recip.AddressEntry._u.PrimarySmtpAddress
        if smtp not in self._app._shared_with_me:
            raise Exception("You do not have permission to perform this operation.")
        return self._app._shared_folders[smtp]

    def GetGlobalAddressList(self):
        app = self._app
        return type("GAL", (), {"AddressEntries": [AddressEntry(u) for u in app._gal]})()


class Store:
    def __init__(self, app, name=None, store_id="STORE-ME", kind=0, root=None):
        self._app = app
        self.DisplayName = name or app._me.PrimarySmtpAddress
        self.StoreID = store_id
        self.ExchangeStoreType = kind
        self._root = root

    def GetRootFolder(self):
        return self._root or self._app._root

    def GetDefaultFolder(self, kind):
        if self._root is not None:
            raise Exception("The default folder is not available in this store.")
        return self._app._folders[kind]


class Application:
    def __init__(self):
        self._me = ExchangeUser("Ori Harel", "ori.harel@contoso.com", "Engineer", "R&D", "TLV")
        self._gal = [
            self._me,
            ExchangeUser("Dana Levi", "dana.levi@contoso.com", "Product Manager", "Product", "TLV"),
            ExchangeUser("Dana Cohen", "dcohen@contoso.com", "Designer", "Design", "Haifa"),
            ExchangeUser("Yossi Mizrahi", "yossi@contoso.com", "Team Lead", "R&D", "TLV"),
            ExchangeUser("Room Galil", "room.galil@contoso.com", "Meeting room", "Facilities", "TLV"),
        ]
        self._by_id: dict = {}
        self._deleted: list = []
        self._outbox: list = []
        self._filters: list = []
        self._occurrences: dict = {}
        self._calendars: dict = {}
        self._store = Store(self)
        self._root = Folder(self, self._me.PrimarySmtpAddress, None)
        self._folders = {}
        for kind, name, dtype in [
            (6, "Inbox", 0), (5, "Sent Items", 0), (16, "Drafts", 0), (3, "Deleted Items", 0),
            (4, "Outbox", 0), (23, "Junk Email", 0), (9, "Calendar", 1), (10, "Contacts", 2),
            (13, "Tasks", 3), (12, "Notes", 5),
        ]:
            f = Folder(self, name, kind, dtype)
            f.Parent = self._root
            self._root._subs.append(f)
            self._folders[kind] = f
        self._shared_with_me = {"dana.levi@contoso.com"}
        self._shared_folders = {}
        self._extra_stores = []
        self.Session = None
        self._ns = Namespace(self)
        self.Session = self._ns
        self._seed()

    # -- address book
    def _directory_lookup(self, text):
        t = text.strip().lower()
        hits = [u for u in self._gal if t in (u.Name.lower(), u.PrimarySmtpAddress.lower(), u.Alias.lower())]
        if not hits:
            hits = [u for u in self._gal if all(w in u.Name.lower() for w in t.split())] if t else []
            if len(hits) != 1:
                hits = hits if len(hits) > 1 else []
        return [AddressEntry(u) for u in hits]

    def _calendar_of(self, smtp):
        if smtp == self._me.PrimarySmtpAddress:
            return [i for i in self._folders[9]._items if i.Class == 26 and i._start]
        return self._calendars.get(smtp, [])

    # -- COM surface
    def GetNamespace(self, name):
        return self._ns

    def CreateItem(self, kind):
        if kind == 0:
            return MailItem(self)
        if kind == 1:
            return AppointmentItem(self)
        if kind == 3:
            return TaskItem(self)
        raise Exception(f"CreateItem({kind}) not simulated")

    # -- seed data (dates relative to "today", local time)
    def _seed(self):
        today = dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        inbox = self._folders[6]

        def mail(subject, sender, hours_ago, html=True, to=None, cc="", unread=False, body=None):
            u = next(g for g in self._gal if g.Name == sender)
            text = body or f"Hi Ori, about {subject}."
            m = MailItem(
                self, Subject=subject, SenderName=u.Name,
                SenderEmailAddress=AddressEntry(u).Address, _sender_smtp=u.PrimarySmtpAddress,
                UnRead=unread, Body=text,
                BodyFormat=2 if html else 1,
                HTMLBody=(f"<html><body><div class=WordSection1><p class=MsoNormal>{text}</p></div></body></html>" if html else ""),
            )
            m._display = {1: to or self._me.Name, 2: cc, 3: ""}
            m._received = dt.datetime.now().replace(second=0, microsecond=0) - dt.timedelta(hours=hours_ago)
            inbox._add(m)
            return m

        # One received 30 minutes after local midnight: a UTC-naive 'since
        # midnight' filter on UTC+2 would miss it.
        early = mail("Early bird", "Yossi Mizrahi", 0, unread=True)
        early._received = today + dt.timedelta(minutes=30)
        mail("Q3 numbers", "Dana Levi", 3, unread=True, cc="Yossi Mizrahi")
        mail("Design review", "Dana Cohen", 26)
        mail("Plain text note", "Yossi Mizrahi", 50, html=False)
        mail("Old thread", "Dana Levi", 24 * 10)

        # A three-message thread: Dana -> my reply (Sent Items) -> Dana.
        first = mail("Offsite plan", "Dana Levi", 30, html=False,
                     body="Can we do the offsite on Thursday?")
        mine = MailItem(self, Subject="RE: Offsite plan", SenderName=self._me.Name,
                        _sender_smtp=self._me.PrimarySmtpAddress, BodyFormat=1, Sent=True,
                        ConversationID=first.ConversationID,
                        Body=("Thursday is fine, booking Room Galil.\n\nOri\n\n"
                              "________________________________\nFrom: Dana Levi\n"
                              "Sent: Thursday\nSubject: Offsite plan\n\n"
                              "Can we do the offsite on Thursday?"))
        mine._display = {1: "Dana Levi", 2: "", 3: ""}
        mine._received = mine._sent_on = first._received + dt.timedelta(hours=1)
        self._folders[5]._add(mine)
        last = mail("RE: Offsite plan", "Dana Levi", 28, html=False,
                    body=("מעולה, תודה!\n\n________________________________\n"
                          "מאת: אורי הראל\nנשלח: יום חמישי\nנושא: RE: Offsite plan\n\n"
                          "Thursday is fine, booking Room Galil."))
        last.ConversationID = first.ConversationID

        # My calendar: tomorrow 10:00-11:00 busy.
        tomorrow = today + dt.timedelta(days=1)
        a = AppointmentItem(self, Subject="Existing 1:1", BusyStatus=2)
        a._start, a._end = tomorrow + dt.timedelta(hours=10), tomorrow + dt.timedelta(hours=11)
        a.Save()

        # A weekly recurring standup (series + 3 occurrences) at 09:00.
        series = AppointmentItem(self, Subject="Standup", RecurrenceState=1, IsRecurring=True,
                                 MeetingStatus=1, BusyStatus=2)
        series._start, series._end = tomorrow + dt.timedelta(hours=9), tomorrow + dt.timedelta(hours=9, minutes=15)
        series.Save()
        occs = []
        for k in range(3):
            o = AppointmentItem(self, Subject="Standup", RecurrenceState=2, IsRecurring=True,
                                MeetingStatus=1, BusyStatus=2)
            o.EntryID = series.EntryID  # occurrences share the series id
            o._start = series._start + dt.timedelta(days=7 * k)
            o._end = series._end + dt.timedelta(days=7 * k)
            o.Parent = self._folders[9]
            occs.append(o)
        self._occurrences[series.EntryID] = occs

        # Colleagues' calendars (free/busy source); Dana Levi's is shared.
        def busy(smtp, day, h0, h1, status=2, subject="Busy"):
            ap = AppointmentItem(self, Subject=subject, BusyStatus=status)
            ap._start = today + dt.timedelta(days=day, hours=h0)
            ap._end = today + dt.timedelta(days=day, hours=h1)
            ap._owner = smtp
            self._calendars.setdefault(smtp, []).append(ap)
            return ap

        busy("dana.levi@contoso.com", 1, 13, 15, subject="Roadmap planning")
        busy("dana.levi@contoso.com", 1, 15, 16, status=1, subject="Maybe: vendor call")
        busy("yossi@contoso.com", 1, 11, 13, subject="Hiring loop")
        # Online archive, a shared team mailbox and a PST.
        def old_mail(subject, sender, days_ago, body=""):
            u = next(g for g in self._gal if g.Name == sender)
            m = MailItem(self, Subject=subject, SenderName=u.Name,
                         SenderEmailAddress=AddressEntry(u).Address, _sender_smtp=u.PrimarySmtpAddress,
                         Body=body or subject, BodyFormat=1)
            m._display = {1: self._me.Name, 2: "", 3: ""}
            m._received = today - dt.timedelta(days=days_ago) + dt.timedelta(hours=10)
            return m

        self._add_store(f"Online Archive - {self._me.PrimarySmtpAddress}", 4, {
            "Inbox": [old_mail("Budget 2024 final", "Dana Levi", 400, "the budget numbers")],
            "Inbox/Projects/Apollo": [old_mail("Apollo kickoff budget", "Yossi Mizrahi", 500)],
            "Deleted Items": [old_mail("Deleted budget draft", "Dana Levi", 450)],
        })
        self._add_store("Support Team", 4, {
            "Inbox": [old_mail("Customer budget question", "Dana Cohen", 2)],
        })
        self._add_store("Old PST 2019", 3, {
            "Personal": [old_mail("Wedding budget", "Yossi Mizrahi", 2000)],
        })

        shared = Folder(self, "Calendar", 9, 1, store_id="STORE-DANA")
        for ap in self._calendars["dana.levi@contoso.com"]:
            shared._add(ap)
        self._shared_folders["dana.levi@contoso.com"] = shared

    def _add_store(self, name, kind, folders):
        """folders: {"Inbox": [mails], "Projects/Apollo": [mails]}"""
        sid = f"STORE-{len(self._extra_stores) + 2}"
        root = Folder(self, name, None, store_id=sid)
        for path, mails in folders.items():
            parent = root
            for seg in path.split("/"):
                match = next((f for f in parent._subs if f.Name == seg), None)
                if match is None:
                    match = Folder(self, seg, None, 0, store_id=sid)
                    match.Parent = parent
                    parent._subs.append(match)
                parent = match
            for m in mails:
                parent._add(m)
        cal = Folder(self, "Calendar", None, 1, store_id=sid)
        root._subs.append(cal)
        self._extra_stores.append(Store(self, name, sid, kind, root))

    def _report(self):
        return {"outbox": self._outbox, "filters": self._filters[-5:]}


_APP = None


def application():
    global _APP
    if _APP is None:
        _APP = Application()
    return _APP


def dump_state():  # handy when debugging a failing e2e test
    return json.dumps(application()._report(), default=str, indent=2)
