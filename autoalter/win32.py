from __future__ import annotations

import ctypes
from ctypes import wintypes

SW_RESTORE = 9
CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002
MOD_NOREPEAT = 0x4000
VK_F2 = 0x71
VK_F3 = 0x72
WM_HOTKEY = 0x0312
PM_REMOVE = 0x0001
HOTKEY_ID_STOP = 1
HOTKEY_ID_START = 2
FAST_POLL_INTERVAL = 0.01
WINDOW_ACTIVATE_DELAY = 0.03
MIN_CLIPBOARD_TIMEOUT = 0.08
CLIPBOARD_EXTRA_DELAY = 0.03

USER32 = ctypes.windll.user32
KERNEL32 = ctypes.windll.kernel32

USER32.OpenClipboard.argtypes = [wintypes.HWND]
USER32.OpenClipboard.restype = wintypes.BOOL
USER32.CloseClipboard.argtypes = []
USER32.CloseClipboard.restype = wintypes.BOOL
USER32.EmptyClipboard.argtypes = []
USER32.EmptyClipboard.restype = wintypes.BOOL
USER32.GetClipboardData.argtypes = [wintypes.UINT]
USER32.GetClipboardData.restype = ctypes.c_void_p
USER32.SetClipboardData.argtypes = [wintypes.UINT, ctypes.c_void_p]
USER32.SetClipboardData.restype = ctypes.c_void_p
USER32.RegisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT]
USER32.RegisterHotKey.restype = wintypes.BOOL
USER32.UnregisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int]
USER32.UnregisterHotKey.restype = wintypes.BOOL
USER32.PeekMessageW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.UINT, wintypes.UINT]
USER32.PeekMessageW.restype = wintypes.BOOL
USER32.GetAsyncKeyState.argtypes = [ctypes.c_int]
USER32.GetAsyncKeyState.restype = ctypes.c_short
KERNEL32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
KERNEL32.GlobalAlloc.restype = ctypes.c_void_p
KERNEL32.GlobalLock.argtypes = [ctypes.c_void_p]
KERNEL32.GlobalLock.restype = ctypes.c_void_p
KERNEL32.GlobalUnlock.argtypes = [ctypes.c_void_p]
KERNEL32.GlobalUnlock.restype = wintypes.BOOL
KERNEL32.GlobalFree.argtypes = [ctypes.c_void_p]
KERNEL32.GlobalFree.restype = ctypes.c_void_p


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class MSG(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("message", wintypes.UINT),
        ("wParam", wintypes.WPARAM),
        ("lParam", wintypes.LPARAM),
        ("time", wintypes.DWORD),
        ("pt", POINT),
    ]


def enable_dpi_awareness() -> None:
    try:
        USER32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except Exception:
        try:
            USER32.SetProcessDPIAware()
        except Exception:
            pass
