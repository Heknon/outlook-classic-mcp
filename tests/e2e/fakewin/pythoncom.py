"""Stand-in for pywin32's pythoncom, for end-to-end tests off Windows."""

COINIT_APARTMENTTHREADED = 2
COINIT_DISABLE_OLE1DDE = 4


class com_error(Exception):  # noqa: N801
    pass


def CoInitializeEx(flags):  # noqa: N802
    pass


def CoInitialize():  # noqa: N802
    pass


def CoUninitialize():  # noqa: N802
    pass


def PumpWaitingMessages():  # noqa: N802
    return 0
