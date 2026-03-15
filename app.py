from __future__ import annotations

import ctypes
import queue
import sys
import threading
import time
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

ROOT = Path(__file__).resolve().parent
VENDOR_DIR = ROOT / ".vendor"
if VENDOR_DIR.exists():
    sys.path.insert(0, str(VENDOR_DIR))

import pyautogui
import tkinter as tk

from autoalter import (
    AppConfig,
    ClipboardManager,
    HOTKEY_ID_START,
    HOTKEY_ID_STOP,
    ItemPoint,
    MOD_NOREPEAT,
    MSG,
    OcrScanner,
    PM_REMOVE,
    PointPicker,
    RegionPicker,
    RelativeRegion,
    TextNormalizer,
    USER32,
    VK_F2,
    VK_F3,
    WM_HOTKEY,
    WindowInfo,
    WindowManager,
    enable_dpi_awareness,
)
from autoalter.config_store import ConfigController
from autoalter.runner import AutomationRunner


CONFIG_PATH = ROOT / "config.json"

CONFIG_PATH = ROOT / "config.json"


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
        self.config_controller = ConfigController(self, CONFIG_PATH)
        self.runner = AutomationRunner(self)
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
        self.current_runtime_config: AppConfig | None = None

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
        self.action_delay_var = tk.StringVar(value="0.12")
        self.click_jitter_var = tk.StringVar(value="2")
        self.human_delay_var = tk.BooleanVar(value=True)
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
        ttk.Label(setup_frame, text="取用等待(秒)").grid(row=6, column=2, sticky="w", pady=(12, 0))
        ttk.Entry(setup_frame, textvariable=self.action_delay_var, width=10).grid(
            row=7, column=2, sticky="ew", padx=(0, 8)
        )
        ttk.Label(setup_frame, text="點擊浮動(px)").grid(row=6, column=3, sticky="w", pady=(12, 0))
        ttk.Entry(setup_frame, textvariable=self.click_jitter_var, width=10).grid(
            row=7, column=3, sticky="ew", padx=(0, 8)
        )
        ttk.Label(setup_frame, text="每輪間隔(秒)").grid(row=6, column=4, sticky="w", pady=(12, 0))
        ttk.Entry(setup_frame, textvariable=self.cycle_delay_var, width=10).grid(
            row=7, column=4, sticky="ew", padx=(0, 8)
        )
        ttk.Checkbutton(
            setup_frame,
            text="人性化延遲(+0.01~0.05秒)",
            variable=self.human_delay_var,
        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Label(
            setup_frame,
            text="F2 可全域立即停止，F3 可全域開始；點擊浮動只會影響左右鍵，不影響 hover 判斷。",
        ).grid(row=8, column=3, columnspan=5, sticky="w", pady=(10, 0))

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
        self.config_controller.load()

    def save_config(self) -> None:
        self.config_controller.save()

    def parse_optional_int(self, raw_value: str) -> int | None:
        return self.config_controller.parse_optional_int(raw_value)

    def collect_config(
        self,
        *,
        require_target: bool = True,
        require_action: bool = False,
        require_items: bool = True,
        require_ocr: bool = True,
    ) -> AppConfig:
        return self.config_controller.collect_config(
            require_target=require_target,
            require_action=require_action,
            require_items=require_items,
            require_ocr=require_ocr,
        )

    def refresh_item_listbox(self) -> None:
        self.config_controller.refresh_item_listbox()

    def match_target_list(self, raw_targets: str, texts: list[str]) -> tuple[bool, str, str]:
        return self.runner.match_target_list(raw_targets, texts)

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
        self.current_runtime_config = None
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
        return self.runner.resolve_window(config, activate_window=activate_window)

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
        return self.runner.absolute_point(window, offset_x, offset_y)

    def absolute_region(self, window: WindowInfo, region: RelativeRegion) -> dict[str, int]:
        return self.runner.absolute_region(window, region)

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
        return self.runner.copy_item_text(config)

    def run_detection_check(
        self, config: AppConfig, window: WindowInfo, item: ItemPoint, stage: str
    ) -> tuple[bool, str, str]:
        return self.runner.run_detection_check(config, window, item, stage)

    def perform_item_action(self, config: AppConfig, window: WindowInfo, item: ItemPoint) -> bool:
        return self.runner.perform_item_action(config, window, item)

    def wait_with_pause(self, seconds: float) -> bool:
        return self.runner.wait_with_pause(seconds)

    def pause_due_to_match(self, item: ItemPoint, target: str, hit: str) -> None:
        self.runner.pause_due_to_match(item, target, hit)

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
        self.current_runtime_config = config
        self.release_shift_for_loop()
        self.shift_action_primed = False
        self.worker = threading.Thread(target=self.run_automation, args=(config,), daemon=True)
        self.worker.start()
        self.set_status("\u57f7\u884c\u4e2d")
        self.append_log(
            f"自動流程已開始，判斷方式={self.detection_mode_label(config.detection_mode)}，取用等待={config.action_delay}s，點擊浮動={config.click_jitter}px。"
        )
        if config.human_delay:
            self.append_log("已啟用人性化延遲：每次等待會額外隨機增加 0.01~0.05 秒。")
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
        self.runner.run_automation(config)

    def on_close(self) -> None:
        self.stop_event.set()
        self.current_runtime_config = None
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








