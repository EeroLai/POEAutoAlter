from __future__ import annotations

import ctypes
from ctypes import wintypes
import json
import queue
import random
import re
import sys
import threading
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

ROOT = Path(__file__).resolve().parent
VENDOR_DIR = ROOT / ".vendor"
if VENDOR_DIR.exists():
    sys.path.insert(0, str(VENDOR_DIR))

import cv2
import mss
import numpy as np
import pyautogui
import tkinter as tk
from opencc import OpenCC
from rapidocr_onnxruntime import RapidOCR


CONFIG_PATH = ROOT / "config.json"
OCR_SCORE_THRESHOLD = 0.25
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


@dataclass
class ItemPoint:
    name: str
    x: int
    y: int

    @classmethod
    def from_dict(cls, data: dict) -> "ItemPoint":
        return cls(
            name=str(data.get("name", "物品")),
            x=int(data.get("x", 0)),
            y=int(data.get("y", 0)),
        )


@dataclass
class RelativeRegion:
    left: int = 0
    top: int = 0
    width: int = 0
    height: int = 0

    @property
    def has_area(self) -> bool:
        return self.width > 0 and self.height > 0

    @classmethod
    def from_dict(cls, data: dict | None) -> "RelativeRegion":
        data = data or {}
        return cls(
            left=int(data.get("left", 0)),
            top=int(data.get("top", 0)),
            width=int(data.get("width", 0)),
            height=int(data.get("height", 0)),
        )


@dataclass
class WindowInfo:
    hwnd: int
    title: str
    left: int
    top: int
    width: int
    height: int


@dataclass
class AnchorInfo:
    left: int
    top: int
    width: int
    height: int
    score: float
    window: WindowInfo


@dataclass
class AppConfig:
    window_title: str = ""
    detection_mode: str = "ocr"
    hover_delay: float = 0.35
    click_delay: float = 0.25
    click_jitter: int = 2
    hold_shift_loop: bool = False
    cycle_delay: float = 0.8
    target_text: str = ""
    action_mode: str = "window"
    action_x: int | None = None
    action_y: int | None = None
    ocr_region: RelativeRegion = field(default_factory=RelativeRegion)
    item_points: list[ItemPoint] = field(default_factory=list)

    @classmethod
    def from_dict(cls, data: dict) -> "AppConfig":
        return cls(
            window_title=str(data.get("window_title", "")),
            detection_mode=str(data.get("detection_mode", "ocr")),
            hover_delay=float(data.get("hover_delay", 0.35)),
            click_delay=float(data.get("click_delay", 0.25)),
            click_jitter=int(data.get("click_jitter", 2)),
            hold_shift_loop=bool(data.get("hold_shift_loop", False)),
            cycle_delay=float(data.get("cycle_delay", 0.8)),
            target_text=str(data.get("target_text", "")),
            action_mode=str(data.get("action_mode", "window")),
            action_x=data.get("action_x"),
            action_y=data.get("action_y"),
            ocr_region=RelativeRegion.from_dict(data.get("ocr_region")),
            item_points=[ItemPoint.from_dict(item) for item in data.get("item_points", [])],
        )


def enable_dpi_awareness() -> None:
    try:
        USER32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except Exception:
        try:
            USER32.SetProcessDPIAware()
        except Exception:
            pass


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", "", text or "").lower()


def parse_target_list(raw_text: str) -> list[str]:
    parts = re.split(r"[\r\n,，;；|]+", raw_text or "")
    results: list[str] = []
    seen: set[str] = set()
    for part in parts:
        candidate = part.strip()
        normalized = normalize_text(candidate)
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        results.append(candidate)
    return results


class TextNormalizer:
    def __init__(self) -> None:
        self.to_traditional = OpenCC("s2t")
        self.to_simplified = OpenCC("t2s")

    def forms(self, text: str) -> set[str]:
        cleaned = normalize_text(text)
        if not cleaned:
            return set()
        forms = {cleaned}
        forms.add(self.to_traditional.convert(cleaned))
        forms.add(self.to_simplified.convert(cleaned))
        return {item for item in forms if item}

    def matches(self, target: str, texts: list[str]) -> tuple[bool, str]:
        target_forms = self.forms(target)
        if not target_forms:
            return False, ""

        for text in texts:
            candidate_forms = self.forms(text)
            for candidate in candidate_forms:
                if any(target_form in candidate for target_form in target_forms):
                    return True, text

        merged_forms = self.forms("".join(texts))
        for candidate in merged_forms:
            if any(target_form in candidate for target_form in target_forms):
                return True, "".join(texts)

        return False, ""

    def matches_any(self, targets: list[str], texts: list[str]) -> tuple[bool, str, str]:
        for target in targets:
            matched, hit = self.matches(target, texts)
            if matched:
                return True, target, hit
        return False, "", ""


class WindowManager:
    def __init__(self) -> None:
        self.enum_callback_type = ctypes.WINFUNCTYPE(
            ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p
        )

    def _window_from_handle(self, hwnd: int) -> WindowInfo | None:
        if not USER32.IsWindowVisible(hwnd):
            return None

        length = USER32.GetWindowTextLengthW(hwnd)
        if length <= 0:
            return None

        title_buffer = ctypes.create_unicode_buffer(length + 1)
        USER32.GetWindowTextW(hwnd, title_buffer, length + 1)
        title = title_buffer.value.strip()
        if not title:
            return None

        client_rect = RECT()
        if USER32.GetClientRect(hwnd, ctypes.byref(client_rect)):
            origin = POINT(0, 0)
            if USER32.ClientToScreen(hwnd, ctypes.byref(origin)):
                width = int(client_rect.right - client_rect.left)
                height = int(client_rect.bottom - client_rect.top)
                if width >= 100 and height >= 100:
                    return WindowInfo(
                        hwnd=int(hwnd),
                        title=title,
                        left=int(origin.x),
                        top=int(origin.y),
                        width=width,
                        height=height,
                    )

        rect = RECT()
        if not USER32.GetWindowRect(hwnd, ctypes.byref(rect)):
            return None

        width = int(rect.right - rect.left)
        height = int(rect.bottom - rect.top)
        if width < 100 or height < 100:
            return None

        return WindowInfo(
            hwnd=int(hwnd),
            title=title,
            left=int(rect.left),
            top=int(rect.top),
            width=width,
            height=height,
        )

    def list_windows(self) -> list[WindowInfo]:
        windows: list[WindowInfo] = []

        @self.enum_callback_type
        def callback(hwnd, _lparam):
            window = self._window_from_handle(int(hwnd))
            if window is not None:
                windows.append(window)
            return True

        USER32.EnumWindows(callback, 0)
        return windows

    def get_foreground_window(self) -> WindowInfo | None:
        hwnd = int(USER32.GetForegroundWindow())
        if not hwnd:
            return None
        return self._window_from_handle(hwnd)

    def find_window(self, title_query: str) -> WindowInfo:
        query = title_query.strip().lower()
        if not query:
            raise ValueError("請輸入視窗標題關鍵字。")

        foreground = self.get_foreground_window()
        if foreground and query in foreground.title.lower():
            return foreground

        matches = [window for window in self.list_windows() if query in window.title.lower()]
        if not matches:
            raise RuntimeError(f"找不到標題包含「{title_query}」的視窗。")

        matches.sort(key=lambda item: (len(item.title), item.title.lower()))
        return matches[0]

    def activate(self, hwnd: int) -> None:
        USER32.ShowWindow(hwnd, SW_RESTORE)
        USER32.SetForegroundWindow(hwnd)


class TemplateLocator:
    def __init__(self) -> None:
        self.cache: dict[str, np.ndarray] = {}

    def load_template(self, image_path: str) -> np.ndarray:
        path = str(Path(image_path).expanduser().resolve())
        cached = self.cache.get(path)
        if cached is not None:
            return cached

        data = np.fromfile(path, dtype=np.uint8)
        if data.size == 0:
            raise RuntimeError(f"讀不到圖片檔案: {path}")

        image = cv2.imdecode(data, cv2.IMREAD_COLOR)
        if image is None:
            raise RuntimeError(f"無法載入圖片: {path}")

        self.cache[path] = image
        return image

    def match_template(
        self, screenshot: np.ndarray, template: np.ndarray
    ) -> tuple[tuple[int, int], float]:
        source_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
        template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

        if (
            source_gray.shape[0] < template_gray.shape[0]
            or source_gray.shape[1] < template_gray.shape[1]
        ):
            raise RuntimeError("按鈕圖片比視窗截圖還大，無法比對。")

        result = cv2.matchTemplate(source_gray, template_gray, cv2.TM_CCOEFF_NORMED)
        _min_val, max_val, _min_loc, max_loc = cv2.minMaxLoc(result)
        return (int(max_loc[0]), int(max_loc[1])), float(max_val)

    def locate(self, window: WindowInfo, image_path: str, threshold: float) -> AnchorInfo:
        template = self.load_template(image_path)
        monitor = {
            "left": window.left,
            "top": window.top,
            "width": window.width,
            "height": window.height,
        }

        with mss.mss() as sct:
            screenshot = np.array(sct.grab(monitor))

        screenshot_bgr = cv2.cvtColor(screenshot, cv2.COLOR_BGRA2BGR)
        (match_x, match_y), score = self.match_template(screenshot_bgr, template)
        if score < threshold:
            raise RuntimeError(f"找不到按鈕圖片，最高相似度只有 {score:.3f}。")

        return AnchorInfo(
            left=window.left + match_x,
            top=window.top + match_y,
            width=int(template.shape[1]),
            height=int(template.shape[0]),
            score=score,
            window=window,
        )


class ClipboardManager:
    def open_clipboard(self) -> None:
        for _ in range(10):
            if USER32.OpenClipboard(0):
                return
            time.sleep(0.03)
        raise RuntimeError("無法開啟剪貼簿。")

    def get_text(self) -> str:
        self.open_clipboard()
        try:
            handle = USER32.GetClipboardData(CF_UNICODETEXT)
            if not handle:
                return ""
            locked = KERNEL32.GlobalLock(handle)
            if not locked:
                return ""
            try:
                return ctypes.wstring_at(locked)
            finally:
                KERNEL32.GlobalUnlock(handle)
        finally:
            USER32.CloseClipboard()

    def set_text(self, text: str) -> None:
        buffer = ctypes.create_unicode_buffer(text)
        handle = KERNEL32.GlobalAlloc(GMEM_MOVEABLE, ctypes.sizeof(buffer))
        if not handle:
            raise RuntimeError("無法配置剪貼簿記憶體。")

        locked = KERNEL32.GlobalLock(handle)
        if not locked:
            KERNEL32.GlobalFree(handle)
            raise RuntimeError("無法鎖定剪貼簿記憶體。")
        try:
            ctypes.memmove(locked, buffer, ctypes.sizeof(buffer))
        finally:
            KERNEL32.GlobalUnlock(handle)

        self.open_clipboard()
        try:
            USER32.EmptyClipboard()
            if not USER32.SetClipboardData(CF_UNICODETEXT, handle):
                KERNEL32.GlobalFree(handle)
                raise RuntimeError("無法寫入剪貼簿。")
        finally:
            USER32.CloseClipboard()


class OcrScanner:
    def __init__(self) -> None:
        self.reader = RapidOCR()

    def scan_monitor(self, monitor: dict[str, int]) -> list[str]:
        if monitor["width"] <= 0 or monitor["height"] <= 0:
            raise ValueError("OCR 區域尺寸必須大於 0。")

        with mss.mss() as sct:
            grab = np.array(sct.grab(monitor))

        source = cv2.cvtColor(grab, cv2.COLOR_BGRA2BGR)
        enlarged = cv2.resize(source, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)
        gray = cv2.cvtColor(enlarged, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (3, 3), 0)
        _threshold, binary = cv2.threshold(
            gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU
        )
        variants = [enlarged, cv2.cvtColor(binary, cv2.COLOR_GRAY2BGR)]

        results: list[str] = []
        seen: set[str] = set()
        for image in variants:
            result, _ = self.reader(image)
            if not result:
                continue
            for item in result:
                _box, text, score = item
                if float(score) < OCR_SCORE_THRESHOLD:
                    continue
                normalized = normalize_text(str(text))
                if not normalized or normalized in seen:
                    continue
                seen.add(normalized)
                results.append(str(text))
        return results

class ScreenOverlayBase(tk.Toplevel):
    def __init__(self, master: tk.Tk, title: str) -> None:
        super().__init__(master)
        with mss.mss() as sct:
            monitor = sct.monitors[0]

        self.offset_left = monitor["left"]
        self.offset_top = monitor["top"]
        self.geometry(
            f"{monitor['width']}x{monitor['height']}+{self.offset_left}+{self.offset_top}"
        )
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.attributes("-alpha", 0.28)
        self.configure(bg="#111111")

        self.canvas = tk.Canvas(self, bg="#111111", highlightthickness=0, cursor="crosshair")
        self.canvas.pack(fill="both", expand=True)
        self.canvas.create_text(
            24,
            24,
            anchor="nw",
            fill="#ffffff",
            font=("Microsoft JhengHei UI", 14, "bold"),
            text=title,
        )

        self.bind("<Escape>", lambda _: self.destroy())
        self.focus_force()
        self.grab_set()


class PointPicker(ScreenOverlayBase):
    def __init__(self, master: tk.Tk, title: str, on_pick) -> None:
        super().__init__(master, title)
        self.on_pick = on_pick
        self.canvas.bind("<Button-1>", self.pick)

    def pick(self, event) -> None:
        x = int(event.x_root)
        y = int(event.y_root)
        self.destroy()
        self.on_pick(x, y)


class RegionPicker(ScreenOverlayBase):
    def __init__(self, master: tk.Tk, title: str, on_pick) -> None:
        super().__init__(master, title)
        self.on_pick = on_pick
        self.start_x = 0
        self.start_y = 0
        self.preview_id = None
        self.canvas.bind("<ButtonPress-1>", self.start_drag)
        self.canvas.bind("<B1-Motion>", self.update_drag)
        self.canvas.bind("<ButtonRelease-1>", self.finish_drag)

    def start_drag(self, event) -> None:
        self.start_x = event.x_root
        self.start_y = event.y_root
        if self.preview_id is not None:
            self.canvas.delete(self.preview_id)
        self.preview_id = self.canvas.create_rectangle(
            event.x,
            event.y,
            event.x,
            event.y,
            outline="#20d5ff",
            width=2,
            dash=(4, 4),
        )

    def update_drag(self, event) -> None:
        if self.preview_id is None:
            return
        self.canvas.coords(
            self.preview_id,
            self.start_x - self.offset_left,
            self.start_y - self.offset_top,
            event.x_root - self.offset_left,
            event.y_root - self.offset_top,
        )

    def finish_drag(self, event) -> None:
        left = min(self.start_x, event.x_root)
        top = min(self.start_y, event.y_root)
        width = abs(event.x_root - self.start_x)
        height = abs(event.y_root - self.start_y)
        self.destroy()
        if width < 5 or height < 5:
            return
        self.on_pick(left, top, width, height)


class AutomationApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        enable_dpi_awareness()
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0

        self.title("AutoAlter 視窗定位與 OCR 自動操作")
        self.geometry("980x860")
        self.minsize(980, 760)

        self.normalizer = TextNormalizer()
        self.window_manager = WindowManager()
        self.clipboard = ClipboardManager()
        self.scanner = OcrScanner()
        self.queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.stop_event = threading.Event()
        self.hotkey_registered = False
        self.start_hotkey_registered = False
        self.hotkey_monitor_stop = threading.Event()
        self.hotkey_monitor: threading.Thread | None = None
        self.shift_loop_key_down = False
        self.shift_action_primed = False
        self.item_points: list[ItemPoint] = []

        self.window_title_var = tk.StringVar(value="Path of Exile")
        self.detection_mode_var = tk.StringVar(value="ocr")
        self.target_text_var = tk.StringVar()
        self.action_x_var = tk.StringVar()
        self.action_y_var = tk.StringVar()
        self.region_left_var = tk.StringVar(value="0")
        self.region_top_var = tk.StringVar(value="0")
        self.region_width_var = tk.StringVar(value="0")
        self.region_height_var = tk.StringVar(value="0")
        self.hover_delay_var = tk.StringVar(value="0.35")
        self.click_delay_var = tk.StringVar(value="0.25")
        self.click_jitter_var = tk.StringVar(value="2")
        self.shift_loop_var = tk.BooleanVar(value=False)
        self.cycle_delay_var = tk.StringVar(value="0.8")
        self.status_var = tk.StringVar(value="待命")
        self.anchor_status_var = tk.StringVar(value="尚未測試視窗")

        self.build_ui()
        self.register_global_hotkeys()
        self.start_hotkey_monitor()
        self.load_config()
        self.after(120, self.process_queue)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        main = ttk.Frame(self, padding=16)
        main.grid(sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(3, weight=1)

        window_frame = ttk.LabelFrame(main, text="視窗設定", padding=14)
        window_frame.grid(row=0, column=0, sticky="ew")
        for index in range(6):
            window_frame.columnconfigure(index, weight=1)

        ttk.Label(window_frame, text="視窗標題關鍵字").grid(row=0, column=0, sticky="w")
        ttk.Entry(window_frame, textvariable=self.window_title_var).grid(
            row=1, column=0, columnspan=4, sticky="ew", padx=(0, 8)
        )
        ttk.Button(window_frame, text="抓前景視窗", command=self.use_foreground_window).grid(
            row=1, column=4, sticky="ew", padx=(0, 8)
        )
        ttk.Button(window_frame, text="測試視窗", command=self.test_window_lookup).grid(
            row=1, column=5, sticky="ew"
        )

        ttk.Label(window_frame, text="視窗狀態").grid(row=2, column=0, sticky="w", pady=(12, 0))
        ttk.Label(
            window_frame,
            textvariable=self.anchor_status_var,
            font=("Microsoft JhengHei UI", 10, "bold"),
        ).grid(row=3, column=0, columnspan=6, sticky="w")
        ttk.Label(
            window_frame,
            text="右鍵點、物品點、OCR 區域都直接相對於 Path of Exile 視窗左上角。",
        ).grid(row=4, column=0, columnspan=6, sticky="w", pady=(10, 0))

        setup_frame = ttk.LabelFrame(main, text="操作設定", padding=14)
        setup_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        for index in range(8):
            setup_frame.columnconfigure(index, weight=1)

        ttk.Label(setup_frame, text="目標文字清單(逗號/分號/| 分隔)").grid(row=0, column=0, sticky="w")
        ttk.Entry(setup_frame, textvariable=self.target_text_var).grid(
            row=1, column=0, columnspan=4, sticky="ew", padx=(0, 8)
        )
        ttk.Label(setup_frame, text="右鍵點 X(選填)").grid(row=0, column=4, sticky="w")
        ttk.Entry(setup_frame, textvariable=self.action_x_var, width=10).grid(
            row=1, column=4, sticky="ew", padx=(0, 8)
        )
        ttk.Label(setup_frame, text="\u53f3\u9375\u9ede Y(\u9078\u586b)").grid(row=0, column=5, sticky="w")
        ttk.Entry(setup_frame, textvariable=self.action_y_var, width=10).grid(
            row=1, column=5, sticky="ew", padx=(0, 8)
        )
        ttk.Button(setup_frame, text="\u6293\u53f3\u9375\u9ede(\u9078\u586b)", command=self.pick_action_point).grid(
            row=1, column=6, columnspan=2, sticky="ew"
        )

        ttk.Checkbutton(
            setup_frame,
            text="\u6574\u6bb5\u5faa\u74b0\u6309\u4f4f Shift",
            variable=self.shift_loop_var,
        ).grid(row=2, column=4, columnspan=4, sticky="w", pady=(12, 0))

        ttk.Label(setup_frame, text="判斷方式").grid(row=2, column=0, sticky="w", pady=(12, 0))
        ttk.Radiobutton(
            setup_frame,
            text="OCR",
            value="ocr",
            variable=self.detection_mode_var,
        ).grid(row=3, column=0, sticky="w")
        ttk.Radiobutton(
            setup_frame,
            text="剪貼簿 Ctrl+C",
            value="clipboard",
            variable=self.detection_mode_var,
        ).grid(row=3, column=1, columnspan=2, sticky="w")
        ttk.Label(
            setup_frame,
            text="剪貼簿模式會 hover 物品後送出 Ctrl+C；OCR 模式才會用到 OCR 區域。",
        ).grid(row=3, column=3, columnspan=5, sticky="w")

        ttk.Label(setup_frame, text="OCR Left").grid(row=4, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(setup_frame, textvariable=self.region_left_var, width=10).grid(
            row=5, column=0, sticky="ew", padx=(0, 8)
        )
        ttk.Label(setup_frame, text="OCR Top").grid(row=4, column=1, sticky="w", pady=(12, 0))
        ttk.Entry(setup_frame, textvariable=self.region_top_var, width=10).grid(
            row=5, column=1, sticky="ew", padx=(0, 8)
        )
        ttk.Label(setup_frame, text="OCR Width").grid(row=4, column=2, sticky="w", pady=(12, 0))
        ttk.Entry(setup_frame, textvariable=self.region_width_var, width=10).grid(
            row=5, column=2, sticky="ew", padx=(0, 8)
        )
        ttk.Label(setup_frame, text="OCR Height").grid(row=4, column=3, sticky="w", pady=(12, 0))
        ttk.Entry(setup_frame, textvariable=self.region_height_var, width=10).grid(
            row=5, column=3, sticky="ew", padx=(0, 8)
        )
        ttk.Button(setup_frame, text="框選 OCR 區域", command=self.pick_ocr_region).grid(
            row=5, column=4, sticky="ew", padx=(0, 8)
        )
        ttk.Button(setup_frame, text="清除 OCR 區域", command=self.clear_region).grid(
            row=5, column=5, sticky="ew", padx=(0, 8)
        )
        ttk.Button(setup_frame, text="測試判斷", command=self.test_ocr).grid(
            row=5, column=6, columnspan=2, sticky="ew"
        )

        ttk.Label(setup_frame, text="hover 等待(秒)").grid(row=6, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(setup_frame, textvariable=self.hover_delay_var, width=10).grid(
            row=7, column=0, sticky="ew", padx=(0, 8)
        )
        ttk.Label(setup_frame, text="點擊間隔(秒)").grid(row=6, column=1, sticky="w", pady=(12, 0))
        ttk.Entry(setup_frame, textvariable=self.click_delay_var, width=10).grid(
            row=7, column=1, sticky="ew", padx=(0, 8)
        )
        ttk.Label(setup_frame, text="點擊浮動(px)").grid(row=6, column=2, sticky="w", pady=(12, 0))
        ttk.Entry(setup_frame, textvariable=self.click_jitter_var, width=10).grid(
            row=7, column=2, sticky="ew", padx=(0, 8)
        )
        ttk.Label(setup_frame, text="每輪間隔(秒)").grid(row=6, column=3, sticky="w", pady=(12, 0))
        ttk.Entry(setup_frame, textvariable=self.cycle_delay_var, width=10).grid(
            row=7, column=3, sticky="ew", padx=(0, 8)
        )
        ttk.Label(
            setup_frame,
            text="F2 \u53ef\u5168\u57df\u7acb\u5373\u505c\u6b62\uff0cF3 \u53ef\u5168\u57df\u958b\u59cb\uff1b\u9ede\u64ca\u6d6e\u52d5\u53ea\u6703\u5f71\u97ff\u5de6\u53f3\u9375\uff0c\u4e0d\u5f71\u97ff hover \u5224\u65b7\u3002",
        ).grid(row=7, column=4, columnspan=4, sticky="w")

        items_frame = ttk.LabelFrame(main, text="物品點位", padding=14)
        items_frame.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        items_frame.columnconfigure(0, weight=1)
        items_frame.columnconfigure(1, weight=0)
        items_frame.rowconfigure(0, weight=1)

        self.item_listbox = tk.Listbox(
            items_frame,
            height=8,
            exportselection=False,
            font=("Consolas", 10),
        )
        self.item_listbox.grid(row=0, column=0, rowspan=4, sticky="nsew", padx=(0, 12))

        ttk.Button(items_frame, text="新增物品點", command=self.add_item_point).grid(
            row=0, column=1, sticky="ew"
        )
        ttk.Button(items_frame, text="刪除選取", command=self.remove_selected_item).grid(
            row=1, column=1, sticky="ew", pady=(8, 0)
        )
        ttk.Button(items_frame, text="清空物品點", command=self.clear_item_points).grid(
            row=2, column=1, sticky="ew", pady=(8, 0)
        )
        ttk.Label(
            items_frame,
            text="新增物品點時，位置會記成相對於 Path of Exile 視窗左上角的座標。",
        ).grid(row=3, column=1, sticky="w", pady=(8, 0))

        log_frame = ttk.LabelFrame(main, text="狀態與執行紀錄", padding=14)
        log_frame.grid(row=3, column=0, sticky="nsew", pady=(12, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=1)

        status_row = ttk.Frame(log_frame)
        status_row.grid(row=0, column=0, sticky="ew")
        status_row.columnconfigure(1, weight=1)
        ttk.Label(status_row, text="目前狀態").grid(row=0, column=0, sticky="w")
        ttk.Label(
            status_row,
            textvariable=self.status_var,
            font=("Microsoft JhengHei UI", 11, "bold"),
        ).grid(row=0, column=1, sticky="w", padx=(12, 0))

        self.log_widget = scrolledtext.ScrolledText(
            log_frame,
            wrap="word",
            font=("Consolas", 10),
            state="disabled",
            height=16,
        )
        self.log_widget.grid(row=1, column=0, sticky="nsew", pady=(12, 0))

        controls = ttk.Frame(main)
        controls.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        for index in range(2):
            controls.columnconfigure(index, weight=1)

        ttk.Button(controls, text="\u958b\u59cb", command=self.start_automation).grid(
            row=0, column=0, sticky="ew", padx=(0, 8)
        )
        ttk.Button(controls, text="\u505c\u6b62", command=self.stop_automation).grid(
            row=0, column=1, sticky="ew"
        )

    def use_foreground_window(self) -> None:
        window = self.window_manager.get_foreground_window()
        if window is None:
            messagebox.showerror("找不到視窗", "目前沒有可用的前景視窗。")
            return
        self.window_title_var.set(window.title)
        self.append_log(f"已帶入前景視窗: {window.title}")

    def test_window_lookup(self) -> None:
        try:
            config = self.collect_config(
                require_target=False,
                require_action=False,
                require_items=False,
                require_ocr=False,
            )
            window = self.window_manager.find_window(config.window_title)
            self.anchor_status_var.set(
                f"視窗: {window.title} ({window.left}, {window.top}, {window.width}x{window.height})"
            )
            self.append_log(f"視窗定位成功: {window.title}")
        except Exception as exc:
            messagebox.showerror("視窗測試失敗", str(exc))

    def load_config(self) -> None:
        if not CONFIG_PATH.exists():
            return
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            config = AppConfig.from_dict(data)
        except Exception:
            return

        legacy_anchor_x = data.get("anchor_x")
        legacy_anchor_y = data.get("anchor_y")
        if legacy_anchor_x is not None and legacy_anchor_y is not None:
            config.item_points = [
                ItemPoint(name=item.name, x=item.x + int(legacy_anchor_x), y=item.y + int(legacy_anchor_y))
                for item in config.item_points
            ]
            if config.item_points:
                self.append_log("已將舊版物品點設定轉換為視窗相對座標。")

        if config.action_mode == "anchor" and config.action_x is not None and config.action_y is not None:
            if legacy_anchor_x is not None and legacy_anchor_y is not None:
                config.action_x += int(legacy_anchor_x)
                config.action_y += int(legacy_anchor_y)
                config.action_mode = "window"
                self.append_log("已將舊版定位點設定轉換為視窗相對座標。")

        if config.window_title:
            self.window_title_var.set(config.window_title)
        self.detection_mode_var.set(config.detection_mode if config.detection_mode in {"ocr", "clipboard"} else "ocr")
        self.target_text_var.set(config.target_text)
        self.action_x_var.set("" if config.action_x is None else str(config.action_x))
        self.action_y_var.set("" if config.action_y is None else str(config.action_y))
        self.region_left_var.set(str(config.ocr_region.left))
        self.region_top_var.set(str(config.ocr_region.top))
        self.region_width_var.set(str(config.ocr_region.width))
        self.region_height_var.set(str(config.ocr_region.height))
        self.hover_delay_var.set(str(config.hover_delay))
        self.click_delay_var.set(str(config.click_delay))
        self.click_jitter_var.set(str(config.click_jitter))
        self.shift_loop_var.set(config.hold_shift_loop)
        self.cycle_delay_var.set(str(config.cycle_delay))
        self.item_points = list(config.item_points)
        self.refresh_item_listbox()

    def save_config(self) -> None:
        try:
            config = self.collect_config(
                require_target=False,
                require_action=False,
                require_items=False,
                require_ocr=False,
            )
        except Exception:
            return

        CONFIG_PATH.write_text(
            json.dumps(asdict(config), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def parse_optional_int(self, raw_value: str) -> int | None:
        value = raw_value.strip()
        if not value:
            return None
        return int(value)

    def collect_config(
        self,
        *,
        require_target: bool = True,
        require_action: bool = False,
        require_items: bool = True,
        require_ocr: bool = True,
    ) -> AppConfig:
        try:
            config = AppConfig(
                window_title=self.window_title_var.get().strip(),
                detection_mode=self.detection_mode_var.get().strip() or "ocr",
                hover_delay=float(self.hover_delay_var.get()),
                click_delay=float(self.click_delay_var.get()),
                click_jitter=int(self.click_jitter_var.get()),
                hold_shift_loop=bool(self.shift_loop_var.get()),
                cycle_delay=float(self.cycle_delay_var.get()),
                target_text=self.target_text_var.get().strip(),
                action_mode="window",
                action_x=self.parse_optional_int(self.action_x_var.get()),
                action_y=self.parse_optional_int(self.action_y_var.get()),
                ocr_region=RelativeRegion(
                    left=int(self.region_left_var.get() or 0),
                    top=int(self.region_top_var.get() or 0),
                    width=int(self.region_width_var.get() or 0),
                    height=int(self.region_height_var.get() or 0),
                ),
                item_points=list(self.item_points),
            )
        except ValueError as exc:
            raise ValueError("數值欄位格式不正確。") from exc

        if not config.window_title:
            raise ValueError("請輸入視窗標題關鍵字。")
        if config.detection_mode not in {"ocr", "clipboard"}:
            raise ValueError("判斷方式不正確。")
        if config.hover_delay < 0:
            raise ValueError("hover 等待時間不能小於 0。")
        if config.click_delay < 0:
            raise ValueError("點擊間隔不能小於 0。")
        if config.click_jitter < 0:
            raise ValueError("點擊浮動不能小於 0。")
        if config.cycle_delay < 0:
            raise ValueError("每輪間隔不能小於 0。")
        if require_target and not parse_target_list(config.target_text):
            raise ValueError("請輸入至少一個要判斷的目標文字。")
        if config.action_mode != "window":
            raise ValueError("定位點模式不正確。")
        if require_action and (config.action_x is None or config.action_y is None):
            raise ValueError("請設定定位點。")
        if require_items and not config.item_points:
            raise ValueError("請至少新增一個物品點。")
        if config.ocr_region.width < 0 or config.ocr_region.height < 0:
            raise ValueError("OCR 區域寬高不能是負數。")
        if config.detection_mode == "ocr" and require_ocr and not config.ocr_region.has_area:
            raise ValueError("OCR 模式需要設定 OCR 區域。")

        return config

    def refresh_item_listbox(self) -> None:
        self.item_listbox.delete(0, "end")
        for item in self.item_points:
            self.item_listbox.insert("end", f"{item.name:<8} window=({item.x}, {item.y})")

    def match_target_list(self, raw_targets: str, texts: list[str]) -> tuple[bool, str, str]:
        targets = parse_target_list(raw_targets)
        return self.normalizer.matches_any(targets, texts)

    def set_status(self, message: str) -> None:
        self.status_var.set(message)

    def append_log(self, message: str) -> None:
        timestamp = time.strftime("%H:%M:%S")
        self.log_widget.configure(state="normal")
        self.log_widget.insert("end", f"[{timestamp}] {message}\n")
        self.log_widget.see("end")
        self.log_widget.configure(state="disabled")

    def register_global_hotkeys(self) -> None:
        try:
            self.hotkey_registered = bool(
                USER32.RegisterHotKey(None, HOTKEY_ID_STOP, MOD_NOREPEAT, VK_F2)
            )
        except Exception:
            self.hotkey_registered = False

        try:
            self.start_hotkey_registered = bool(
                USER32.RegisterHotKey(None, HOTKEY_ID_START, MOD_NOREPEAT, VK_F3)
            )
        except Exception:
            self.start_hotkey_registered = False

        registered: list[str] = []
        if self.hotkey_registered:
            registered.append("F2=停止")
        if self.start_hotkey_registered:
            registered.append("F3=\u958b\u59cb")

        if registered:
            self.append_log(f"已註冊全域快捷鍵：{'，'.join(registered)}。")
        if not self.hotkey_registered:
            self.append_log("註冊全域快捷鍵 F2 失敗，請改用停止按鈕。")
        if not self.start_hotkey_registered:
            self.append_log("\u8a3b\u518a\u5168\u57df\u5feb\u6377\u9375 F3 \u5931\u6557\uff0c\u8acb\u6539\u7528\u958b\u59cb\u6309\u9215\u3002")

    def unregister_global_hotkeys(self) -> None:
        if self.hotkey_registered:
            USER32.UnregisterHotKey(None, HOTKEY_ID_STOP)
            self.hotkey_registered = False
        if self.start_hotkey_registered:
            USER32.UnregisterHotKey(None, HOTKEY_ID_START)
            self.start_hotkey_registered = False

    def start_hotkey_monitor(self) -> None:
        self.hotkey_monitor_stop.clear()
        self.hotkey_monitor = threading.Thread(
            target=self.monitor_stop_hotkey_loop,
            name="HotkeyMonitor",
            daemon=True,
        )
        self.hotkey_monitor.start()

    def monitor_stop_hotkey_loop(self) -> None:
        was_f2_pressed = False
        was_f3_pressed = False
        while not self.hotkey_monitor_stop.is_set():
            try:
                is_f2_pressed = bool(USER32.GetAsyncKeyState(VK_F2) & 0x8000)
                is_f3_pressed = bool(USER32.GetAsyncKeyState(VK_F3) & 0x8000)
            except Exception:
                return

            if is_f2_pressed and not was_f2_pressed:
                self.request_stop("F2 全域停止已觸發。")
            if is_f3_pressed and not was_f3_pressed:
                self.request_start("F3 \u5168\u57df\u958b\u59cb\u5df2\u89f8\u767c\u3002")

            was_f2_pressed = is_f2_pressed
            was_f3_pressed = is_f3_pressed
            time.sleep(0.03)

    def request_stop(self, reason: str) -> bool:
        worker = self.worker
        if not worker or not worker.is_alive():
            return False
        if self.stop_event.is_set():
            return False
        self.stop_event.set()
        self.release_shift_for_loop()
        self.shift_action_primed = False
        self.queue_status("\u505c\u6b62\u4e2d")
        self.queue_log(reason)
        return True

    def request_start(self, reason: str) -> bool:
        worker = self.worker
        if worker and worker.is_alive():
            return False
        self.queue.put(("command", "start"))
        self.queue_log(reason)
        return True

    def handle_global_stop_hotkey(self) -> None:
        self.request_stop("F2 全域停止已觸發。")

    def handle_global_start_hotkey(self) -> None:
        self.request_start("F3 \u5168\u57df\u958b\u59cb\u5df2\u89f8\u767c\u3002")

    def process_hotkeys(self) -> None:
        if not (self.hotkey_registered or self.start_hotkey_registered):
            return
        message = MSG()
        while USER32.PeekMessageW(ctypes.byref(message), None, WM_HOTKEY, WM_HOTKEY, PM_REMOVE):
            if message.message != WM_HOTKEY:
                continue
            hotkey_id = int(message.wParam)
            if hotkey_id == HOTKEY_ID_STOP:
                self.handle_global_stop_hotkey()
            elif hotkey_id == HOTKEY_ID_START:
                self.handle_global_start_hotkey()

    def process_queue(self) -> None:
        self.process_hotkeys()
        while True:
            try:
                kind, payload = self.queue.get_nowait()
            except queue.Empty:
                break

            if kind == "status":
                self.set_status(payload)
            elif kind == "log":
                self.append_log(payload)
            elif kind == "anchor":
                self.anchor_status_var.set(payload)
            elif kind == "command" and payload == "start":
                self.start_automation()

        self.after(120, self.process_queue)

    def queue_status(self, message: str) -> None:
        self.queue.put(("status", message))

    def queue_log(self, message: str) -> None:
        self.queue.put(("log", message))

    def queue_anchor(self, message: str) -> None:
        self.queue.put(("anchor", message))

    def with_hidden_window(self, picker_builder) -> None:
        self.withdraw()

        def restore() -> None:
            if not self.winfo_exists():
                return
            self.deiconify()
            self.lift()
            self.focus_force()

        def launch() -> None:
            try:
                picker = picker_builder()
            except Exception as exc:
                restore()
                messagebox.showerror("操作失敗", str(exc))
                return
            if picker is None:
                restore()
                return
            picker.bind("<Destroy>", lambda _: self.after(60, restore), add="+")

        self.after(180, launch)

    def resolve_window(self, config: AppConfig, activate_window: bool = True) -> WindowInfo:
        window = self.window_manager.find_window(config.window_title)
        if activate_window:
            foreground = USER32.GetForegroundWindow()
            if int(foreground or 0) != window.hwnd:
                self.window_manager.activate(window.hwnd)
                time.sleep(WINDOW_ACTIVATE_DELAY)
        self.queue_anchor(
            f"{window.title} | window=({window.left}, {window.top}, {window.width}x{window.height})"
        )
        return window

    def pick_action_point(self) -> None:
        try:
            config = self.collect_config(
                require_target=False,
                require_action=False,
                require_items=False,
                require_ocr=False,
            )
        except Exception as exc:
            messagebox.showerror("設定有誤", str(exc))
            return

        self.with_hidden_window(
            lambda: PointPicker(
                self,
                "點一下定位點，會記成相對於 Path of Exile 視窗左上角的座標",
                lambda x, y: self.set_action_point_from_screen(config, x, y),
            )
        )

    def set_action_point_from_screen(self, config: AppConfig, x: int, y: int) -> None:
        window = self.window_manager.find_window(config.window_title)
        offset_x = x - window.left
        offset_y = y - window.top
        self.action_x_var.set(str(offset_x))
        self.action_y_var.set(str(offset_y))
        self.append_log(
            f"定位點已更新: absolute=({x}, {y}), window_offset=({offset_x}, {offset_y})"
        )

    def pick_ocr_region(self) -> None:
        try:
            config = self.collect_config(
                require_target=False,
                require_action=False,
                require_items=False,
                require_ocr=False,
            )
        except Exception as exc:
            messagebox.showerror("設定有誤", str(exc))
            return

        self.with_hidden_window(
            lambda: RegionPicker(
                self,
                "拖曳框出 OCR 區域，會記成相對於 Path of Exile 視窗左上角的座標",
                lambda left, top, width, height: self.set_region_from_screen(
                    config, left, top, width, height
                ),
            )
        )

    def set_region_from_screen(
        self, config: AppConfig, left: int, top: int, width: int, height: int
    ) -> None:
        window = self.window_manager.find_window(config.window_title)
        offset_left = left - window.left
        offset_top = top - window.top
        self.region_left_var.set(str(offset_left))
        self.region_top_var.set(str(offset_top))
        self.region_width_var.set(str(width))
        self.region_height_var.set(str(height))
        self.append_log(
            "OCR 區域已更新: "
            f"absolute=({left}, {top}, {width}, {height}), "
            f"window_offset=({offset_left}, {offset_top}, {width}, {height})"
        )

    def clear_region(self) -> None:
        self.region_left_var.set("0")
        self.region_top_var.set("0")
        self.region_width_var.set("0")
        self.region_height_var.set("0")
        self.append_log("已清除 OCR 區域。")

    def add_item_point(self) -> None:
        try:
            config = self.collect_config(
                require_target=False,
                require_action=False,
                require_items=False,
                require_ocr=False,
            )
        except Exception as exc:
            messagebox.showerror("設定有誤", str(exc))
            return

        self.with_hidden_window(
            lambda: PointPicker(
                self,
                "點一下物品位置，會記成相對於 Path of Exile 視窗左上角的座標",
                lambda x, y: self.add_item_point_from_screen(config, x, y),
            )
        )

    def add_item_point_from_screen(self, config: AppConfig, x: int, y: int) -> None:
        window = self.resolve_window(config, activate_window=False)
        offset_x = x - window.left
        offset_y = y - window.top
        item = ItemPoint(name=f"Item{len(self.item_points) + 1}", x=offset_x, y=offset_y)
        self.item_points.append(item)
        self.refresh_item_listbox()
        self.append_log(
            f"已新增 {item.name}: absolute=({x}, {y}), window_offset=({offset_x}, {offset_y})"
        )

    def remove_selected_item(self) -> None:
        selection = self.item_listbox.curselection()
        if not selection:
            return
        index = int(selection[0])
        removed = self.item_points.pop(index)
        self.refresh_item_listbox()
        self.append_log(f"已刪除物品點: {removed.name}")

    def clear_item_points(self) -> None:
        self.item_points.clear()
        self.refresh_item_listbox()
        self.append_log("已清空所有物品點。")

    def absolute_point(self, window: WindowInfo, offset_x: int, offset_y: int) -> tuple[int, int]:
        return window.left + offset_x, window.top + offset_y

    def absolute_region(self, window: WindowInfo, region: RelativeRegion) -> dict[str, int]:
        return {
            "left": window.left + region.left,
            "top": window.top + region.top,
            "width": region.width,
            "height": region.height,
        }

    def press_shift_for_loop(self) -> None:
        if self.shift_loop_key_down:
            return
        pyautogui.keyDown("shift")
        self.shift_loop_key_down = True

    def release_shift_for_loop(self) -> None:
        if not self.shift_loop_key_down:
            return
        try:
            pyautogui.keyUp("shift")
        finally:
            self.shift_loop_key_down = False

    def sync_shift_for_loop(self, config: AppConfig) -> None:
        should_hold = (
            config.hold_shift_loop
            and self.shift_action_primed
            and not self.stop_event.is_set()
        )
        if should_hold:
            self.press_shift_for_loop()
        else:
            self.release_shift_for_loop()

    def jittered_point(self, x: int, y: int, jitter: int) -> tuple[int, int]:
        if jitter <= 0:
            return x, y
        return x + random.randint(-jitter, jitter), y + random.randint(-jitter, jitter)

    def click_point(self, x: int, y: int, button: str, jitter: int = 0) -> tuple[int, int]:
        click_x, click_y = self.jittered_point(x, y, jitter)
        pyautogui.moveTo(click_x, click_y, duration=0)
        pyautogui.mouseDown(button=button)
        pyautogui.mouseUp(button=button)
        return click_x, click_y

    def detection_mode_label(self, mode: str) -> str:
        return "剪貼簿 Ctrl+C" if mode == "clipboard" else "OCR"

    def copy_item_text(self, config: AppConfig) -> str:
        sentinel = f"__AUTOALTER__{time.time_ns()}__"
        self.clipboard.set_text(sentinel)
        pyautogui.keyDown("ctrl")
        pyautogui.press("c")
        pyautogui.keyUp("ctrl")
        timeout = time.perf_counter() + max(MIN_CLIPBOARD_TIMEOUT, config.click_delay + CLIPBOARD_EXTRA_DELAY)
        while time.perf_counter() < timeout:
            if self.stop_event.is_set():
                return ""
            copied = self.clipboard.get_text().strip()
            if copied and copied != sentinel:
                return copied
            time.sleep(FAST_POLL_INTERVAL)
        return ""

    def run_detection_check(
        self, config: AppConfig, window: WindowInfo, item: ItemPoint, stage: str
    ) -> tuple[bool, str, str]:
        item_x, item_y = self.absolute_point(window, item.x, item.y)
        pyautogui.moveTo(item_x, item_y, duration=0)
        self.queue_log(f"{stage} hover {item.name}: ({item_x}, {item_y})")
        if self.wait_with_pause(config.hover_delay):
            return False, "", ""

        texts: list[str] = []
        if config.detection_mode == "clipboard":
            copied = self.copy_item_text(config)
            if copied:
                texts = [copied]
                self.queue_log(f"{stage} 剪貼簿[{item.name}]: {copied[:160]}")
            else:
                self.queue_log(f"{stage} 剪貼簿[{item.name}]: 沒抓到內容")
        else:
            region = self.absolute_region(window, config.ocr_region)
            texts = self.scanner.scan_monitor(region)
            if texts:
                self.queue_log(f"{stage} OCR[{item.name}]: {' | '.join(texts)}")
            else:
                self.queue_log(f"{stage} OCR[{item.name}]: 沒抓到文字")

        return self.match_target_list(config.target_text, texts)

    def perform_item_action(self, config: AppConfig, window: WindowInfo, item: ItemPoint) -> bool:
        item_x, item_y = self.absolute_point(window, item.x, item.y)

        if config.hold_shift_loop:
            if config.action_x is None or config.action_y is None:
                return False

            action_x = window.left + config.action_x
            action_y = window.top + config.action_y

            if not self.shift_action_primed:
                self.release_shift_for_loop()
                actual_action_x, actual_action_y = self.click_point(
                    action_x,
                    action_y,
                    button="right",
                    jitter=config.click_jitter,
                )
                self.shift_action_primed = True
                self.queue_log(
                    f"\u6574\u6bb5 Shift \u6a21\u5f0f\u5df2\u53d6\u7528\u5b9a\u4f4d\u9ede: ({actual_action_x}, {actual_action_y}) | base=({action_x}, {action_y})"
                )
                if self.wait_with_pause(config.click_delay):
                    return True

            self.sync_shift_for_loop(config)
            actual_item_x, actual_item_y = self.click_point(
                item_x,
                item_y,
                button="left",
                jitter=config.click_jitter,
            )
            self.queue_log(
                f"\u5c0d {item.name} \u6309 Shift+\u5de6\u9375: ({actual_item_x}, {actual_item_y}) | base=({item_x}, {item_y})"
            )
            return self.wait_with_pause(config.click_delay)

        if config.action_x is None or config.action_y is None:
            return False

        action_x = window.left + config.action_x
        action_y = window.top + config.action_y

        actual_action_x, actual_action_y = self.click_point(
            action_x,
            action_y,
            button="right",
            jitter=config.click_jitter,
        )
        self.queue_log(
            f"\u5c0d\u5b9a\u4f4d\u9ede\u6309\u53f3\u9375: ({actual_action_x}, {actual_action_y}) | base=({action_x}, {action_y})"
        )
        if self.wait_with_pause(config.click_delay):
            return True

        actual_item_x, actual_item_y = self.click_point(
            item_x,
            item_y,
            button="left",
            jitter=config.click_jitter,
        )
        self.queue_log(
            f"\u5c0d {item.name} \u6309\u5de6\u9375: ({actual_item_x}, {actual_item_y}) | base=({item_x}, {item_y})"
        )
        return self.wait_with_pause(config.click_delay)

    def wait_with_pause(self, seconds: float) -> bool:
        end_time = time.perf_counter() + seconds
        while time.perf_counter() < end_time:
            if self.stop_event.is_set():
                self.release_shift_for_loop()
                return True
            if self.stop_event.is_set():
                self.release_shift_for_loop()
                return True
            remaining = end_time - time.perf_counter()
            if remaining <= 0:
                break
            time.sleep(min(FAST_POLL_INTERVAL, remaining))
        return False

    def pause_due_to_match(self, item: ItemPoint, target: str, hit: str) -> None:
        self.release_shift_for_loop()
        self.stop_event.set()
        self.queue_status("\u547d\u4e2d\u5f8c\u505c\u6b62")
        self.queue_log(f"{item.name} \u547d\u4e2d\u76ee\u6a19\u6587\u5b57 [{target}]\uff0c\u5df2\u505c\u6b62: {hit}")

    def test_ocr(self) -> None:
        try:
            config = self.collect_config(
                require_target=False,
                require_action=False,
                require_items=False,
                require_ocr=False,
            )
            window = self.resolve_window(config)
        except Exception as exc:
            messagebox.showerror("判斷測試失敗", str(exc))
            return

        if config.detection_mode == "clipboard":
            if not config.item_points:
                messagebox.showerror("判斷測試失敗", "剪貼簿模式至少需要一個物品點。")
                return
            selection = self.item_listbox.curselection()
            index = int(selection[0]) if selection else 0
            item = config.item_points[index]
            matched, target, hit = self.run_detection_check(config, window, item, "測試")
            if matched:
                self.set_status("測試命中")
                self.append_log(f"剪貼簿測試命中 [{target}]: {hit}")
            else:
                self.set_status("測試完成")
                self.append_log("剪貼簿測試未命中目標文字。")
            return

        try:
            region = self.absolute_region(window, config.ocr_region)
            self.append_log(
                f"OCR 測試區域: left={region['left']}, top={region['top']}, width={region['width']}, height={region['height']}"
            )
            texts = self.scanner.scan_monitor(region)
        except Exception as exc:
            messagebox.showerror("OCR 測試失敗", str(exc))
            return

        if texts:
            matched, target, hit = self.match_target_list(config.target_text, texts)
            self.append_log(f"OCR 測試結果: {' | '.join(texts)}")
            if matched:
                self.set_status("測試命中")
                self.append_log(f"OCR 測試命中 [{target}]: {hit}")
            else:
                self.set_status("測試完成")
                self.append_log("OCR 測試未命中目標文字。")
        else:
            self.set_status("測試完成")
            self.append_log("OCR 測試沒抓到文字。")
    def start_automation(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("\u57f7\u884c\u4e2d", "\u76ee\u524d\u6d41\u7a0b\u5df2\u7d93\u5728\u57f7\u884c\u3002")
            return

        try:
            config = self.collect_config(require_action=False)
        except Exception as exc:
            messagebox.showerror("\u8a2d\u5b9a\u6709\u8aa4", str(exc))
            return

        self.save_config()
        self.stop_event.clear()
        self.release_shift_for_loop()
        self.shift_action_primed = False
        self.worker = threading.Thread(target=self.run_automation, args=(config,), daemon=True)
        self.worker.start()
        self.set_status("\u57f7\u884c\u4e2d")
        self.append_log(
            f"\u81ea\u52d5\u6d41\u7a0b\u5df2\u958b\u59cb\uff0c\u5224\u65b7\u65b9\u5f0f={self.detection_mode_label(config.detection_mode)}\uff0c\u9ede\u64ca\u6d6e\u52d5={config.click_jitter}px\u3002"
        )
        if self.hotkey_registered or self.start_hotkey_registered:
            self.append_log("F2 \u53ef\u5168\u57df\u7acb\u5373\u505c\u6b62\uff0cF3 \u53ef\u5168\u57df\u958b\u59cb\u3002")
        if config.action_x is None or config.action_y is None:
            self.append_log("\u672a\u8a2d\u5b9a\u53f3\u9375\u9ede\uff0c\u5c07\u53ea\u505a hover + \u5224\u65b7\u3002")
        if config.hold_shift_loop:
            self.append_log("\u6574\u6bb5 Shift \u6a21\u5f0f\u5df2\u555f\u7528\uff1a\u9996\u6b21\u53d6\u7528\u5f8c\uff0c\u5faa\u74b0\u671f\u9593\u6703\u6301\u7e8c\u6309\u4f4f Shift\u3002")

    def stop_automation(self, reason: str = "停止訊號已送出。") -> None:
        if not self.request_stop(reason):
            self.set_status("待命")

    def run_automation(self, config: AppConfig) -> None:
        cycle = 0
        try:
            while not self.stop_event.is_set():
                self.sync_shift_for_loop(config)
                cycle += 1
                self.queue_status(f"\u5de1\u6aa2\u4e2d\uff0c\u7b2c {cycle} \u8f2a")
                window = self.resolve_window(config)
                self.queue_log(
                    f"\u7b2c {cycle} \u8f2a\u958b\u59cb\uff0c\u8996\u7a97=({window.left}, {window.top}, {window.width}x{window.height})"
                )

                matched_this_cycle = False
                for item in config.item_points:
                    if self.stop_event.is_set():
                        break

                    self.sync_shift_for_loop(config)
                    matched, target, hit = self.run_detection_check(config, window, item, "\u64cd\u4f5c\u524d")
                    if matched:
                        self.pause_due_to_match(item, target, hit)
                        matched_this_cycle = True
                        break
                    if self.stop_event.is_set():
                        break

                    if self.perform_item_action(config, window, item):
                        break

                    self.sync_shift_for_loop(config)
                    matched, target, hit = self.run_detection_check(config, window, item, "\u64cd\u4f5c\u5f8c")
                    if matched:
                        self.pause_due_to_match(item, target, hit)
                        matched_this_cycle = True
                        break
                    if self.stop_event.is_set():
                        break

                if self.stop_event.is_set():
                    break
                if matched_this_cycle:
                    continue
                if self.wait_with_pause(config.cycle_delay):
                    break

            self.queue_status("\u5df2\u505c\u6b62")
            self.queue_log("\u81ea\u52d5\u6d41\u7a0b\u5df2\u7d50\u675f\u3002")
        except pyautogui.FailSafeException:
            self.queue_status("\u5b89\u5168\u66ab\u505c")
            self.queue_log("\u6ed1\u9f20\u79fb\u5230\u5de6\u4e0a\u89d2\u89f8\u767c PyAutoGUI failsafe\uff0c\u6d41\u7a0b\u5df2\u505c\u6b62\u3002")
        except Exception as exc:
            self.queue_status("\u57f7\u884c\u5931\u6557")
            self.queue_log(f"\u81ea\u52d5\u6d41\u7a0b\u5931\u6557: {exc}")
        finally:
            self.release_shift_for_loop()
            self.shift_action_primed = False

    def on_close(self) -> None:
        self.stop_event.set()
        self.release_shift_for_loop()
        self.shift_action_primed = False
        self.hotkey_monitor_stop.set()
        self.unregister_global_hotkeys()
        self.save_config()
        self.destroy()

def main() -> None:
    app = AutomationApp()
    app.mainloop()


if __name__ == "__main__":
    main()








