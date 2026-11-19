"""Gives the program's window a sharp icon in the Windows taskbar.

The screen opens in an Edge "app" window. Edge builds that window's icon
from the page's small 16-pixel picture and stretches it, which looks blurry
in the taskbar. Here we hand the window our own icon, loaded at exactly the
size its screen needs (from app.ico, which holds 16 to 256 pixel versions).
Edge may put its own icon back when the page reloads, so this is re-checked
every second. Everything is wrapped so that a problem here can never stop
the program.
"""
from __future__ import annotations

import ctypes
import os
import sys

TITLES = {"Outlook Archiver", "أرشفة Outlook"}  # the page title in each language
BROWSERS = {"msedge.exe", "chrome.exe"}

WM_GETICON, WM_SETICON = 0x007F, 0x0080
ICON_SMALL, ICON_BIG = 0, 1
IMAGE_ICON, LR_LOADFROMFILE = 1, 0x0010
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000


class WindowIcon:
    def __init__(self, ico_path: str) -> None:
        self.path = ico_path
        self.ok = sys.platform == "win32" and os.path.isfile(ico_path)
        self._icons: dict[int, int] = {}
        if not self.ok:
            return
        try:
            self._setup()
        except Exception:  # noqa: BLE001
            self.ok = False

    def _setup(self) -> None:
        from ctypes import wintypes as wt

        u = ctypes.WinDLL("user32", use_last_error=True)
        k = ctypes.WinDLL("kernel32", use_last_error=True)
        self.WNDENUMPROC = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)
        u.EnumWindows.argtypes = [self.WNDENUMPROC, wt.LPARAM]
        u.IsWindowVisible.argtypes = [wt.HWND]
        u.GetWindowTextLengthW.argtypes = [wt.HWND]
        u.GetWindowTextW.argtypes = [wt.HWND, wt.LPWSTR, ctypes.c_int]
        u.GetWindowThreadProcessId.argtypes = [wt.HWND, ctypes.POINTER(wt.DWORD)]
        u.SendMessageW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM]
        u.SendMessageW.restype = ctypes.c_ssize_t
        u.LoadImageW.argtypes = [wt.HINSTANCE, wt.LPCWSTR, wt.UINT, ctypes.c_int, ctypes.c_int, wt.UINT]
        u.LoadImageW.restype = wt.HANDLE
        try:
            u.GetDpiForWindow.argtypes = [wt.HWND]
            u.GetDpiForWindow.restype = wt.UINT
            self._dpi = u.GetDpiForWindow
        except AttributeError:  # Windows older than 10 (1607)
            self._dpi = None
        k.OpenProcess.argtypes = [wt.DWORD, wt.BOOL, wt.DWORD]
        k.OpenProcess.restype = wt.HANDLE
        k.QueryFullProcessImageNameW.argtypes = [wt.HANDLE, wt.DWORD, wt.LPWSTR, ctypes.POINTER(wt.DWORD)]
        k.CloseHandle.argtypes = [wt.HANDLE]
        self.u, self.k, self.wt = u, k, wt

    # ------------------------------------------------------------ helpers
    def _icon(self, size: int) -> int:
        if size not in self._icons:
            h = self.u.LoadImageW(None, self.path, IMAGE_ICON, size, size, LR_LOADFROMFILE)
            self._icons[size] = h or 0  # kept for the program's lifetime: the window uses it
        return self._icons[size]

    def _process_name(self, hwnd) -> str:
        pid = self.wt.DWORD()
        self.u.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        h = self.k.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid.value)
        if not h:
            return ""
        try:
            buf = ctypes.create_unicode_buffer(1024)
            size = self.wt.DWORD(1024)
            if not self.k.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
                return ""
            return os.path.basename(buf.value).lower()
        finally:
            self.k.CloseHandle(h)

    def _our_windows(self) -> list:
        found = []

        def check(hwnd, _lparam):
            if self.u.IsWindowVisible(hwnd):
                n = self.u.GetWindowTextLengthW(hwnd)
                if 0 < n < 100:
                    buf = ctypes.create_unicode_buffer(n + 1)
                    self.u.GetWindowTextW(hwnd, buf, n + 1)
                    if buf.value in TITLES and self._process_name(hwnd) in BROWSERS:
                        found.append(hwnd)
            return True

        self.u.EnumWindows(self.WNDENUMPROC(check), 0)
        return found

    # --------------------------------------------------------------- main
    def apply(self) -> int:
        """Set our icon on the program's window(s). Returns how many were found."""
        if not self.ok:
            return 0
        try:
            windows = self._our_windows()
            for hwnd in windows:
                dpi = (self._dpi(hwnd) if self._dpi else 0) or 96
                for which, base in ((ICON_BIG, 32), (ICON_SMALL, 16)):
                    icon = self._icon(round(base * dpi / 96))
                    if icon and self.u.SendMessageW(hwnd, WM_GETICON, which, 0) != icon:
                        self.u.SendMessageW(hwnd, WM_SETICON, which, icon)
            return len(windows)
        except Exception:  # noqa: BLE001 - never let the icon stop the program
            return 0
