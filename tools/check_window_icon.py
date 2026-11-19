"""Build check (Windows): is the program's window really showing our icon?

Run while OutlookArchiver.exe is open. Finds its Edge window, draws the
window's taskbar icon and our own icon at the same size, and compares them
pixel by pixel. Also saves a screenshot to outlook-archiver-window-check.png.
"""
import ctypes
import sys
import time
from ctypes import wintypes as wt

sys.path.insert(0, ".")
from archiver.win_icon import ICON_BIG, ICON_SMALL, WM_GETICON, WindowIcon  # noqa: E402

gdi = ctypes.WinDLL("gdi32")
user = ctypes.WinDLL("user32")
gdi.CreateCompatibleDC.restype = wt.HDC
gdi.CreateCompatibleDC.argtypes = [wt.HDC]
gdi.CreateDIBSection.restype = wt.HBITMAP
gdi.CreateDIBSection.argtypes = [wt.HDC, ctypes.c_void_p, wt.UINT, ctypes.POINTER(ctypes.c_void_p), wt.HANDLE, wt.DWORD]
gdi.SelectObject.argtypes = [wt.HDC, wt.HGDIOBJ]
gdi.DeleteObject.argtypes = [wt.HGDIOBJ]
gdi.DeleteDC.argtypes = [wt.HDC]
user.DrawIconEx.argtypes = [wt.HDC, ctypes.c_int, ctypes.c_int, wt.HANDLE, ctypes.c_int, ctypes.c_int, wt.UINT, wt.HBRUSH, wt.UINT]


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", wt.DWORD), ("biWidth", wt.LONG), ("biHeight", wt.LONG), ("biPlanes", wt.WORD),
                ("biBitCount", wt.WORD), ("biCompression", wt.DWORD), ("biSizeImage", wt.DWORD),
                ("biXPelsPerMeter", wt.LONG), ("biYPelsPerMeter", wt.LONG), ("biClrUsed", wt.DWORD),
                ("biClrImportant", wt.DWORD)]


def pixels(hicon: int, size: int) -> bytes:
    dc = gdi.CreateCompatibleDC(None)
    bih = BITMAPINFOHEADER(ctypes.sizeof(BITMAPINFOHEADER), size, -size, 1, 32, 0, 0, 0, 0, 0, 0)
    bits = ctypes.c_void_p()
    bmp = gdi.CreateDIBSection(dc, ctypes.byref(bih), 0, ctypes.byref(bits), None, 0)
    old = gdi.SelectObject(dc, bmp)
    user.DrawIconEx(dc, 0, 0, hicon, size, size, 0, None, 3)
    data = ctypes.string_at(bits, size * size * 4)
    gdi.SelectObject(dc, old)
    gdi.DeleteObject(bmp)
    gdi.DeleteDC(dc)
    return data


def main() -> int:
    wi = WindowIcon("archiver/web/app.ico")
    windows = []
    for _ in range(60):
        windows = wi._our_windows()
        if windows:
            break
        time.sleep(1)
    print("program windows found:", len(windows))
    if not windows:
        return 1
    hwnd = windows[0]
    time.sleep(4)  # give the program time to set its icon
    dpi = (wi._dpi(hwnd) if wi._dpi else 96) or 96
    ok = True
    for which, base, name in ((ICON_BIG, 32, "taskbar (big)"), (ICON_SMALL, 16, "title bar (small)")):
        size = round(base * dpi / 96)
        theirs = wi.u.SendMessageW(hwnd, WM_GETICON, which, 0)
        if not theirs:
            print(name, ": the window has no icon")
            ok = False
            continue
        a, b = pixels(theirs, size), pixels(wi._icon(size), size)
        diff = sum(abs(x - y) for x, y in zip(a, b)) / len(a)
        print(f"{name}: {size}px at {dpi} dpi, mean pixel difference from our icon = {diff:.2f}")
        ok = ok and diff < 3
    try:
        from PIL import ImageGrab
        ImageGrab.grab().save("outlook-archiver-window-check.png")
    except Exception as e:  # noqa: BLE001
        print("screenshot failed:", e)
    print("RESULT:", "our sharp icon is on the window" if ok else "icon NOT applied")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
