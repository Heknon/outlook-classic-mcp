"""Small in-memory stand-ins for Outlook COM objects (for unit tests)."""

from __future__ import annotations

from types import SimpleNamespace


class FakeRecipient:
    def __init__(self, address, type_=1, resolved=True):
        self.Address = address
        self.Name = address
        self.Type = type_
        self.Resolved = resolved


class FakeRecipients:
    def __init__(self, items=None):
        self._items = list(items or [])
        self.unresolvable: set[str] = set()

    def Add(self, address):
        r = FakeRecipient(address, resolved=address not in self.unresolvable)
        self._items.append(r)
        return r

    def ResolveAll(self):
        return all(r.Resolved for r in self._items)

    def Item(self, i):
        return self._items[i - 1]

    def Remove(self, i):
        del self._items[i - 1]

    @property
    def Count(self):
        return len(self._items)

    def __iter__(self):
        return iter(list(self._items))


class FakePropertyAccessor:
    def __init__(self):
        self.props = {}

    def SetProperty(self, name, value):
        self.props[name] = value


class FakeAttachments:
    def __init__(self):
        self.added = []

    def Add(self, path, *args):
        att = SimpleNamespace(path=path, args=args, PropertyAccessor=FakePropertyAccessor())
        self.added.append(att)
        return att

    @property
    def Count(self):
        return len(self.added)

    def __iter__(self):
        return iter(self.added)


class FakeMail:
    Class = 43

    def __init__(self, *, html=None, body="", entry_id="NEW1", body_format=None):
        self.HTMLBody = html or ""
        self.Body = body
        self.BodyFormat = body_format if body_format is not None else (2 if html else 1)
        self.Subject = "RE: Topic"
        self.To = ""
        self.CC = ""
        self.BCC = ""
        self.Importance = 1
        self.EntryID = entry_id
        self.Recipients = FakeRecipients()
        self.Attachments = FakeAttachments()
        self.SendUsingAccount = None
        self.Sent = False
        self.saved = 0
        self.sent = 0

    def Save(self):
        self.saved += 1

    def Send(self):
        self.sent += 1


class FakeOriginal:
    def __init__(self, reply):
        self._reply = reply
        self.calls = []

    def Reply(self):
        self.calls.append("Reply")
        return self._reply

    def ReplyAll(self):
        self.calls.append("ReplyAll")
        return self._reply

    def Forward(self):
        self.calls.append("Forward")
        return self._reply


class FakeNamespace:
    def __init__(self, items):
        self.items = items
        self.CurrentProfileName = "Outlook"

    def GetItemFromID(self, entry_id, store_id=None):
        return self.items[entry_id]


class FakePattern:
    def __init__(self, occurrence):
        self.occurrence = occurrence
        self.asked = []

    def GetOccurrence(self, when):
        self.asked.append(when)
        return self.occurrence


class FakeAppt:
    def __init__(self, *, meeting=False, organizer=True, recurrence_state=0, occurrence=None):
        self.Subject = ""
        self.Start = None
        self.End = None
        self.AllDayEvent = False
        self.Location = ""
        self.Body = ""
        self.MeetingStatus = 1 if meeting else 0
        self.ResponseStatus = 1 if organizer else 3
        self.RecurrenceState = recurrence_state
        self.IsRecurring = recurrence_state != 0
        self.ReminderSet = False
        self.ReminderMinutesBeforeStart = 15
        self.BusyStatus = 2
        self.EntryID = "APPT1"
        self.Recipients = FakeRecipients()
        self.events = []
        self._pattern = FakePattern(occurrence)

    def GetRecurrencePattern(self):
        return self._pattern

    def Save(self):
        self.events.append("Save")

    def Send(self):
        self.events.append("Send")

    def Delete(self):
        self.events.append("Delete")


def fake_outlook_creating(item):
    return SimpleNamespace(CreateItem=lambda kind: item)
