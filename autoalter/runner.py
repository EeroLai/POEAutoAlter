from __future__ import annotations

import random
import time
from typing import Any

import pyautogui

from .models import AppConfig, ItemPoint, RelativeRegion, WindowInfo
from .text_utils import parse_target_list
from .win32 import CLIPBOARD_EXTRA_DELAY, FAST_POLL_INTERVAL, MIN_CLIPBOARD_TIMEOUT, USER32, WINDOW_ACTIVATE_DELAY

HUMAN_DELAY_MIN = 0.01
HUMAN_DELAY_MAX = 0.05


class AutomationRunner:
    def __init__(self, app: Any) -> None:
        self.app = app

    def match_target_list(self, raw_targets: str, texts: list[str]) -> tuple[bool, str, str]:
        targets = parse_target_list(raw_targets)
        return self.app.normalizer.matches_any(targets, texts)

    def resolve_window(self, config: AppConfig, activate_window: bool = True) -> WindowInfo:
        window = self.app.window_manager.find_window(config.window_title)
        if activate_window:
            foreground = USER32.GetForegroundWindow()
            if int(foreground or 0) != window.hwnd:
                self.app.window_manager.activate(window.hwnd)
                time.sleep(WINDOW_ACTIVATE_DELAY)
        self.app.queue_anchor(
            f"{window.title} | window=({window.left}, {window.top}, {window.width}x{window.height})"
        )
        return window

    def absolute_point(self, window: WindowInfo, offset_x: int, offset_y: int) -> tuple[int, int]:
        return window.left + offset_x, window.top + offset_y

    def absolute_region(self, window: WindowInfo, region: RelativeRegion) -> dict[str, int]:
        return {
            "left": window.left + region.left,
            "top": window.top + region.top,
            "width": region.width,
            "height": region.height,
        }

    def copy_item_text(self, config: AppConfig) -> str:
        sentinel = f"__AUTOALTER__{time.time_ns()}__"
        self.app.clipboard.set_text(sentinel)
        pyautogui.keyDown("ctrl")
        pyautogui.press("c")
        pyautogui.keyUp("ctrl")
        timeout = time.perf_counter() + max(MIN_CLIPBOARD_TIMEOUT, config.click_delay + CLIPBOARD_EXTRA_DELAY)
        while time.perf_counter() < timeout:
            if self.app.stop_event.is_set():
                return ""
            copied = self.app.clipboard.get_text().strip()
            if copied and copied != sentinel:
                return copied
            time.sleep(FAST_POLL_INTERVAL)
        return ""

    def run_detection_check(
        self, config: AppConfig, window: WindowInfo, item: ItemPoint, stage: str
    ) -> tuple[bool, str, str]:
        item_x, item_y = self.absolute_point(window, item.x, item.y)
        pyautogui.moveTo(item_x, item_y, duration=0)
        self.app.queue_log(f"{stage} hover {item.name}: ({item_x}, {item_y})")
        if self.wait_with_pause(config.hover_delay):
            return False, "", ""

        texts: list[str] = []
        if config.detection_mode == "clipboard":
            copied = self.copy_item_text(config)
            if copied:
                texts = [copied]
                self.app.queue_log(f"{stage} 剪貼簿[{item.name}]: {copied[:160]}")
            else:
                self.app.queue_log(f"{stage} 剪貼簿[{item.name}]: 沒抓到內容")
        else:
            region = self.absolute_region(window, config.ocr_region)
            texts = self.app.scanner.scan_monitor(region)
            if texts:
                self.app.queue_log(f"{stage} OCR[{item.name}]: {' | '.join(texts)}")
            else:
                self.app.queue_log(f"{stage} OCR[{item.name}]: 沒抓到文字")

        return self.match_target_list(config.target_text, texts)

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

    def perform_item_action(self, config: AppConfig, window: WindowInfo, item: ItemPoint) -> bool:
        item_x, item_y = self.absolute_point(window, item.x, item.y)

        if config.hold_shift_loop:
            if config.action_x is None or config.action_y is None:
                return False

            action_x = window.left + config.action_x
            action_y = window.top + config.action_y

            if not self.app.shift_action_primed:
                self.app.release_shift_for_loop()
                actual_action_x, actual_action_y = self.click_point(
                    action_x,
                    action_y,
                    button="right",
                    jitter=config.click_jitter,
                )
                self.app.shift_action_primed = True
                self.app.queue_log(
                    f"整段 Shift 模式已取用定位點: ({actual_action_x}, {actual_action_y}) | base=({action_x}, {action_y})"
                )
                if self.wait_with_pause(config.action_delay):
                    return True

            self.app.sync_shift_for_loop(config)
            actual_item_x, actual_item_y = self.click_point(
                item_x,
                item_y,
                button="left",
                jitter=config.click_jitter,
            )
            self.app.queue_log(
                f"對 {item.name} 按 Shift+左鍵: ({actual_item_x}, {actual_item_y}) | base=({item_x}, {item_y})"
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
        self.app.queue_log(
            f"對定位點按右鍵: ({actual_action_x}, {actual_action_y}) | base=({action_x}, {action_y})"
        )
        if self.wait_with_pause(config.action_delay):
            return True

        actual_item_x, actual_item_y = self.click_point(
            item_x,
            item_y,
            button="left",
            jitter=config.click_jitter,
        )
        self.app.queue_log(
            f"對 {item.name} 按左鍵: ({actual_item_x}, {actual_item_y}) | base=({item_x}, {item_y})"
        )
        return self.wait_with_pause(config.click_delay)

    def wait_with_pause(self, seconds: float) -> bool:
        if seconds < 0:
            seconds = 0
        config = getattr(self.app, "current_runtime_config", None)
        if config and getattr(config, "human_delay", False):
            seconds += random.uniform(HUMAN_DELAY_MIN, HUMAN_DELAY_MAX)
        end_time = time.perf_counter() + seconds
        while time.perf_counter() < end_time:
            if self.app.stop_event.is_set():
                self.app.release_shift_for_loop()
                return True
            remaining = end_time - time.perf_counter()
            if remaining <= 0:
                break
            time.sleep(min(FAST_POLL_INTERVAL, remaining))
        return False

    def pause_due_to_match(self, item: ItemPoint, target: str, hit: str) -> None:
        self.app.release_shift_for_loop()
        self.app.stop_event.set()
        self.app.queue_status("命中後停止")
        self.app.queue_log(f"{item.name} 命中目標文字 [{target}]，已停止: {hit}")

    def run_automation(self, config: AppConfig) -> None:
        cycle = 0
        try:
            while not self.app.stop_event.is_set():
                self.app.sync_shift_for_loop(config)
                cycle += 1
                self.app.queue_status(f"巡檢中，第 {cycle} 輪")
                window = self.resolve_window(config)
                self.app.queue_log(
                    f"第 {cycle} 輪開始，視窗=({window.left}, {window.top}, {window.width}x{window.height})"
                )

                matched_this_cycle = False
                for item in config.item_points:
                    if self.app.stop_event.is_set():
                        break

                    self.app.sync_shift_for_loop(config)
                    matched, target, hit = self.run_detection_check(config, window, item, "操作前")
                    if matched:
                        self.pause_due_to_match(item, target, hit)
                        matched_this_cycle = True
                        break
                    if self.app.stop_event.is_set():
                        break

                    if self.perform_item_action(config, window, item):
                        break

                    self.app.sync_shift_for_loop(config)
                    matched, target, hit = self.run_detection_check(config, window, item, "操作後")
                    if matched:
                        self.pause_due_to_match(item, target, hit)
                        matched_this_cycle = True
                        break
                    if self.app.stop_event.is_set():
                        break

                if self.app.stop_event.is_set():
                    break
                if matched_this_cycle:
                    continue
                if self.wait_with_pause(config.cycle_delay):
                    break

            self.app.queue_status("已停止")
            self.app.queue_log("自動流程已結束。")
        except pyautogui.FailSafeException:
            self.app.queue_status("安全暫停")
            self.app.queue_log("滑鼠移到左上角觸發 PyAutoGUI failsafe，流程已停止。")
        except Exception as exc:
            self.app.queue_status("執行失敗")
            self.app.queue_log(f"自動流程失敗: {exc}")
        finally:
            self.app.release_shift_for_loop()
            self.app.shift_action_primed = False
