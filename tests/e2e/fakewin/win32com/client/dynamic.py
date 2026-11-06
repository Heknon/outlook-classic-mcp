"""Stand-in for win32com.client.dynamic: Dispatch returns the simulated Outlook."""

import fake_outlook


def Dispatch(name):  # noqa: N802
    assert name == "Outlook.Application", name
    return fake_outlook.application()
