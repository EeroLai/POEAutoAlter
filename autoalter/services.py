from __future__ import annotations

import ctypes
import time
from pathlib import Path

import cv2
import mss
import numpy as np
from rapidocr_onnxruntime import RapidOCR

from .models import AnchorInfo, WindowInfo
from .text_utils import normalize_text
from .win32 import CF_UNICODETEXT, GMEM_MOVEABLE, KERNEL32, SW_RESTORE, USER32, POINT, RECT

OCR_SCORE_THRESHOLD = 0.25


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
            raise ValueError('Please enter a window title keyword.')

        foreground = self.get_foreground_window()
        if foreground and query in foreground.title.lower():
            return foreground

        matches = [window for window in self.list_windows() if query in window.title.lower()]
        if not matches:
            raise RuntimeError(f'No window matched "{title_query}".')

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
            raise RuntimeError(f'Unable to read template image: {path}')

        image = cv2.imdecode(data, cv2.IMREAD_COLOR)
        if image is None:
            raise RuntimeError(f'Unable to decode template image: {path}')

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
            raise RuntimeError('Template image is larger than the captured region.')

        result = cv2.matchTemplate(source_gray, template_gray, cv2.TM_CCOEFF_NORMED)
        _min_val, max_val, _min_loc, max_loc = cv2.minMaxLoc(result)
        return (int(max_loc[0]), int(max_loc[1])), float(max_val)

    def locate(self, window: WindowInfo, image_path: str, threshold: float) -> AnchorInfo:
        template = self.load_template(image_path)
        monitor = {
            'left': window.left,
            'top': window.top,
            'width': window.width,
            'height': window.height,
        }

        with mss.mss() as sct:
            screenshot = np.array(sct.grab(monitor))

        screenshot_bgr = cv2.cvtColor(screenshot, cv2.COLOR_BGRA2BGR)
        (match_x, match_y), score = self.match_template(screenshot_bgr, template)
        if score < threshold:
            raise RuntimeError(f'Template score below threshold: {score:.3f}.')

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
        raise RuntimeError('Unable to open the clipboard.')

    def get_text(self) -> str:
        self.open_clipboard()
        try:
            handle = USER32.GetClipboardData(CF_UNICODETEXT)
            if not handle:
                return ''
            locked = KERNEL32.GlobalLock(handle)
            if not locked:
                return ''
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
            raise RuntimeError('Unable to allocate clipboard memory.')

        locked = KERNEL32.GlobalLock(handle)
        if not locked:
            KERNEL32.GlobalFree(handle)
            raise RuntimeError('Unable to lock clipboard memory.')
        try:
            ctypes.memmove(locked, buffer, ctypes.sizeof(buffer))
        finally:
            KERNEL32.GlobalUnlock(handle)

        self.open_clipboard()
        try:
            USER32.EmptyClipboard()
            if not USER32.SetClipboardData(CF_UNICODETEXT, handle):
                KERNEL32.GlobalFree(handle)
                raise RuntimeError('Unable to write clipboard data.')
        finally:
            USER32.CloseClipboard()


class OcrScanner:
    def __init__(self) -> None:
        self.reader = RapidOCR()

    def scan_monitor(self, monitor: dict[str, int]) -> list[str]:
        if monitor['width'] <= 0 or monitor['height'] <= 0:
            raise ValueError('OCR region width and height must be greater than 0.')

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
