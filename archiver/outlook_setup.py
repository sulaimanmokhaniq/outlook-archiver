"""Which Outlook is on this laptop, and running classic Outlook next to the new one.

Classic Outlook can run hidden in the background while the person keeps
using the new Outlook; this program drives that hidden copy to read mail and
to make PST files (only classic Outlook can make them - the new Outlook
also relies on classic Outlook for its PST features).

The "New Outlook" switch is a per-user setting (UseNewOutlook). If it stops
classic Outlook from starting in the background, the program turns it off
for the few seconds classic Outlook needs to start, then puts it back
exactly as it was. The person's choice of app is never changed.
"""
from __future__ import annotations

from contextlib import contextmanager
from typing import Iterator

PREFS = r"Software\Microsoft\Office\16.0\Outlook\Preferences"
PROFILES = r"Software\Microsoft\Office\16.0\Outlook\Profiles"
APP_PATH = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
SWITCH = "UseNewOutlook"


class Registry:
    """The few registry reads and writes needed (HKCU / HKLM only)."""

    def __init__(self) -> None:
        import winreg

        self.w = winreg
        self.HKCU = winreg.HKEY_CURRENT_USER
        self.HKLM = winreg.HKEY_LOCAL_MACHINE

    def get_dword(self, root, path: str, name: str) -> int | None:
        try:
            with self.w.OpenKey(root, path) as k:
                value, kind = self.w.QueryValueEx(k, name)
                return int(value) if kind == self.w.REG_DWORD else None
        except OSError:
            return None

    def set_dword(self, root, path: str, name: str, value: int) -> None:
        with self.w.CreateKeyEx(root, path, 0, self.w.KEY_SET_VALUE) as k:
            self.w.SetValueEx(k, name, 0, self.w.REG_DWORD, int(value))

    def key_exists(self, root, path: str) -> bool:
        try:
            with self.w.OpenKey(root, path):
                return True
        except OSError:
            return False

    def has_subkeys(self, root, path: str) -> bool:
        try:
            with self.w.OpenKey(root, path) as k:
                return self.w.QueryInfoKey(k)[0] > 0
        except OSError:
            return False


def status(reg=None) -> dict:
    """What is on this laptop. Unknown answers are reported as None."""
    try:
        reg = reg or Registry()
    except ImportError:  # not Windows
        return {"classic_installed": None, "new_outlook_on": False, "has_profile": None}
    return {
        "classic_installed": reg.key_exists(reg.HKLM, APP_PATH) or None,  # missing key is not proof
        "new_outlook_on": reg.get_dword(reg.HKCU, PREFS, SWITCH) == 1,
        "has_profile": reg.has_subkeys(reg.HKCU, PROFILES),
    }


@contextmanager
def classic_can_start(reg=None) -> Iterator[bool]:
    """Briefly turn off the "New Outlook" switch so classic Outlook can start.

    Yields True if the switch was changed. It is always put back, even if
    starting classic Outlook fails.
    """
    reg = reg or Registry()
    before = reg.get_dword(reg.HKCU, PREFS, SWITCH)
    if before != 1:
        yield False
        return
    reg.set_dword(reg.HKCU, PREFS, SWITCH, 0)
    try:
        yield True
    finally:
        reg.set_dword(reg.HKCU, PREFS, SWITCH, before)
