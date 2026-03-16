from __future__ import annotations

import random
import time
from typing import Any

import pyautogui

from .models import AppConfig, ItemPoint, RelativeRegion, WindowInfo
from .text_utils import parse_target_list
from .win32 import (
    CLIPBOARD_EXTRA_DELAY,
    FAST_POLL_INTERVAL,
    MIN_CLIPBOARD_TIMEOUT,
    USER32,
    WINDOW_ACTIVATE_DELAY,
)

HUMAN_DELAY_MIN = 0.01
HUMAN_DELAY_MAX = 0.05
STALE_TEXT_RECHECKS = 3
STALE_TEXT_WAIT = 0.08
REALISTIC_MOVE_MIN = 0.03
REALISTIC_MOVE_MAX = 0.12
REALISTIC_PAUSE_CHANCE = 0.18
REALISTIC_PAUSE_MIN = 0.05
REALISTIC_PAUSE_MAX = 0.18


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
        timeout = time.perf_counter() + max(
            MIN_CLIPBOARD_TIMEOUT,
            config.click_delay + CLIPBOARD_EXTRA_DELAY,
        )
        while time.perf_counter() < timeout:
            if self.app.stop_event.is_set():
                return ""
            copied = self.app.clipboard.get_text().strip()
            if copied and copied != sentinel:
                return copied
            time.sleep(FAST_POLL_INTERVAL)
        return ""

    def capture_item_text(
        self,
        config: AppConfig,
        window: WindowInfo,
        item: ItemPoint,
        stage: str,
    ) -> str:
        item_x, item_y = self.absolute_point(window, item.x, item.y)
        pyautogui.moveTo(item_x, item_y, duration=self.movement_duration(config))
        self.maybe_realistic_pause(config)
        self.app.queue_log(f"{stage} hover {item.name}: ({item_x}, {item_y})")
        if self.wait_with_pause(config.hover_delay):
            return ""

        copied = self.copy_item_text(config)
        if copied:
            self.app.queue_log(f"{stage} \u526a\u8cbc\u7c3f[{item.name}]: {copied[:160]}")
        else:
            self.app.queue_log(f"{stage} \u526a\u8cbc\u7c3f[{item.name}]: \u6c92\u6293\u5230\u5167\u5bb9")
        return copied

    def run_detection_check(
        self,
        config: AppConfig,
        window: WindowInfo,
        item: ItemPoint,
        stage: str,
    ) -> tuple[bool, str, str]:
        copied = self.capture_item_text(config, window, item, stage)
        texts: list[str] = [copied] if copied else []
        return self.match_target_list(config.target_text, texts)

    def jittered_point(self, x: int, y: int, jitter: int) -> tuple[int, int]:
        if jitter <= 0:
            return x, y
        return x + random.randint(-jitter, jitter), y + random.randint(-jitter, jitter)

    def movement_duration(self, config: AppConfig) -> float:
        if getattr(config, "realistic_mode", False):
            return random.uniform(REALISTIC_MOVE_MIN, REALISTIC_MOVE_MAX)
        return 0.0

    def maybe_realistic_pause(self, config: AppConfig) -> bool:
        if not getattr(config, "realistic_mode", False):
            return False
        if random.random() >= REALISTIC_PAUSE_CHANCE:
            return False
        return self.wait_with_pause(random.uniform(REALISTIC_PAUSE_MIN, REALISTIC_PAUSE_MAX))

    def click_point(self, x: int, y: int, button: str, jitter: int = 0, config: AppConfig | None = None) -> tuple[int, int]:
        click_x, click_y = self.jittered_point(x, y, jitter)
        move_duration = self.movement_duration(config) if config is not None else 0.0
        pyautogui.moveTo(click_x, click_y, duration=move_duration)
        pyautogui.mouseDown(button=button)
        pyautogui.mouseUp(button=button)
        if config is not None:
            self.maybe_realistic_pause(config)
        return click_x, click_y

    def detection_mode_label(self, mode: str) -> str:
        return "\u526a\u8cbc\u7c3f Ctrl+C"

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
                    config=config,
                )
                self.app.shift_action_primed = True
                self.app.queue_log(
                    f"\u6574\u6bb5 Shift \u6a21\u5f0f\u5df2\u53d6\u7528\u5b9a\u4f4d\u9ede: ({actual_action_x}, {actual_action_y}) | base=({action_x}, {action_y})"
                )
                if self.wait_with_pause(config.action_delay):
                    return True

            self.app.sync_shift_for_loop(config)
            actual_item_x, actual_item_y = self.click_point(
                item_x,
                item_y,
                button="left",
                jitter=config.click_jitter,
                config=config,
            )
            self.app.queue_log(
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
        self.app.queue_log(
            f"\u5c0d\u5b9a\u4f4d\u9ede\u6309\u53f3\u9375: ({actual_action_x}, {actual_action_y}) | base=({action_x}, {action_y})"
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
            f"\u5c0d {item.name} \u6309\u5de6\u9375: ({actual_item_x}, {actual_item_y}) | base=({item_x}, {item_y})"
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
        self.app.queue_status("\u547d\u4e2d\u5f8c\u505c\u6b62")
        self.app.queue_log(
            f"{item.name} \u547d\u4e2d\u76ee\u6a19\u6587\u5b57 [{target}]\uff0c\u5df2\u505c\u6b62: {hit}"
        )

    def run_automation(self, config: AppConfig) -> None:
        cycle = 1
        can_retry_item = config.action_x is not None and config.action_y is not None
        try:
            self.app.sync_shift_for_loop(config)
            self.app.queue_status(f"\u5de1\u6aa2\u4e2d\uff0c\u7b2c {cycle} \u8f2a")
            window = self.resolve_window(config)
            self.app.queue_log(
                f"\u7b2c {cycle} \u8f2a\u958b\u59cb\uff0c\u8996\u7a97=({window.left}, {window.top}, {window.width}x{window.height})"
            )

            if not can_retry_item:
                self.app.queue_log("\u672a\u8a2d\u5b9a\u6539\u9020\u77f3\u4f4d\u7f6e\uff0c\u5c07\u7121\u6cd5\u5c0d\u55ae\u4e00\u7269\u54c1\u9ede\u6301\u7e8c\u91cd\u8a66\u3002")

            for item in config.item_points:
                if self.app.stop_event.is_set():
                    break

                self.app.queue_log(f"\u958b\u59cb\u8655\u7406 {item.name}\u3002")
                while not self.app.stop_event.is_set():
                    before_text = self.capture_item_text(config, window, item, "\u64cd\u4f5c\u524d")
                    matched, target, hit = self.match_target_list(
                        config.target_text,
                        [before_text] if before_text else [],
                    )
                    if matched:
                        self.app.queue_log(f"{item.name} \u5df2\u547d\u4e2d [{target}]\uff0c\u5207\u63db\u4e0b\u4e00\u500b\u7269\u54c1\u9ede\u3002")
                        break
                    if self.app.stop_event.is_set():
                        break

                    if not can_retry_item:
                        self.app.queue_log(f"{item.name} \u5c1a\u672a\u547d\u4e2d\uff0c\u4e14\u672a\u8a2d\u5b9a\u6539\u9020\u77f3\u4f4d\u7f6e\uff0c\u505c\u6b62\u6d41\u7a0b\u3002")
                        self.app.stop_event.set()
                        break

                    if self.perform_item_action(config, window, item):
                        break

                    stale_detected = False
                    matched = False
                    target = ""
                    hit = ""
                    for recheck_index in range(STALE_TEXT_RECHECKS):
                        self.app.sync_shift_for_loop(config)
                        after_text = self.capture_item_text(config, window, item, "\u64cd\u4f5c\u5f8c")
                        matched, target, hit = self.match_target_list(
                            config.target_text,
                            [after_text] if after_text else [],
                        )
                        if matched:
                            self.app.queue_log(f"{item.name} \u5df2\u547d\u4e2d [{target}]\uff0c\u5207\u63db\u4e0b\u4e00\u500b\u7269\u54c1\u9ede\u3002")
                            stale_detected = False
                            break
                        if self.app.stop_event.is_set():
                            break
                        if before_text and after_text and after_text == before_text:
                            stale_detected = True
                            self.app.queue_log(
                                f"{item.name} \u5167\u5bb9\u5c1a\u672a\u66f4\u65b0\uff0c\u7b49\u5f85\u5f8c\u91cd\u65b0\u78ba\u8a8d ({recheck_index + 1}/{STALE_TEXT_RECHECKS})\u3002"
                            )
                            extra_wait = max(config.action_delay, STALE_TEXT_WAIT)
                            if getattr(config, "realistic_mode", False):
                                extra_wait += random.uniform(0.03, 0.12)
                            if self.wait_with_pause(extra_wait):
                                break
                            continue
                        stale_detected = False
                        break
                    else:
                        stale_detected = True

                    if self.app.stop_event.is_set():
                        break
                    if matched:
                        break
                    if stale_detected:
                        self.app.queue_log(f"{item.name} \u5167\u5bb9\u4ecd\u672a\u66f4\u65b0\uff0c\u672c\u6b21\u4e0d\u8ffd\u52a0\u9ede\u64ca\uff0c\u91cd\u65b0\u9032\u5165\u5224\u65b7\u3002")
                        continue

                if self.app.stop_event.is_set():
                    break

            if not self.app.stop_event.is_set():
                self.app.queue_status("\u5df2\u505c\u6b62")
                self.app.queue_log("\u6240\u6709\u7269\u54c1\u9ede\u90fd\u5df2\u5b8c\u6210\uff0c\u6d41\u7a0b\u5df2\u505c\u6b62\u3002")
        except pyautogui.FailSafeException:
            self.app.queue_status("\u5b89\u5168\u66ab\u505c")
            self.app.queue_log("\u6ed1\u9f20\u79fb\u5230\u5de6\u4e0a\u89d2\u89f8\u767c PyAutoGUI failsafe\uff0c\u6d41\u7a0b\u5df2\u505c\u6b62\u3002")
        except Exception as exc:
            self.app.queue_status("\u57f7\u884c\u5931\u6557")
            self.app.queue_log(f"\u81ea\u52d5\u6d41\u7a0b\u5931\u6557: {exc}")
        finally:
            self.app.release_shift_for_loop()
            self.app.shift_action_primed = False
