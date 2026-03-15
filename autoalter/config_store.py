from __future__ import annotations

import json
from dataclasses import asdict
from pathlib import Path
from typing import Any

from .models import AppConfig, ItemPoint, RelativeRegion
from .text_utils import parse_target_list


class ConfigController:
    def __init__(self, app: Any, config_path: Path) -> None:
        self.app = app
        self.config_path = config_path

    def load(self) -> None:
        if not self.config_path.exists():
            return
        try:
            data = json.loads(self.config_path.read_text(encoding="utf-8"))
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
                self.app.append_log("已將舊版物品點設定轉換為視窗相對座標。")

        if config.action_mode == "anchor" and config.action_x is not None and config.action_y is not None:
            if legacy_anchor_x is not None and legacy_anchor_y is not None:
                config.action_x += int(legacy_anchor_x)
                config.action_y += int(legacy_anchor_y)
                config.action_mode = "window"
                self.app.append_log("已將舊版定位點設定轉換為視窗相對座標。")

        if config.window_title:
            self.app.window_title_var.set(config.window_title)
        self.app.detection_mode_var.set("clipboard")
        self.app.target_text_var.set(config.target_text)
        self.app.action_x_var.set("" if config.action_x is None else str(config.action_x))
        self.app.action_y_var.set("" if config.action_y is None else str(config.action_y))
        self.app.region_left_var.set(str(config.ocr_region.left))
        self.app.region_top_var.set(str(config.ocr_region.top))
        self.app.region_width_var.set(str(config.ocr_region.width))
        self.app.region_height_var.set(str(config.ocr_region.height))
        self.app.hover_delay_var.set(str(config.hover_delay))
        self.app.click_delay_var.set(str(config.click_delay))
        self.app.action_delay_var.set(str(config.action_delay))
        self.app.click_jitter_var.set(str(config.click_jitter))
        self.app.human_delay_var.set(config.human_delay)
        self.app.shift_loop_var.set(config.hold_shift_loop)
        self.app.cycle_delay_var.set(str(config.cycle_delay))
        self.app.item_points = list(config.item_points)
        self.refresh_item_listbox()

    def save(self) -> None:
        try:
            config = self.collect_config(
                require_target=False,
                require_action=False,
                require_items=False,
                require_ocr=False,
            )
        except Exception:
            return

        self.config_path.write_text(
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
                window_title=self.app.window_title_var.get().strip(),
                detection_mode="clipboard",
                hover_delay=float(self.app.hover_delay_var.get()),
                click_delay=float(self.app.click_delay_var.get()),
                action_delay=float(self.app.action_delay_var.get()),
                click_jitter=int(self.app.click_jitter_var.get()),
                human_delay=bool(self.app.human_delay_var.get()),
                hold_shift_loop=bool(self.app.shift_loop_var.get()),
                cycle_delay=float(self.app.cycle_delay_var.get()),
                target_text=self.app.target_text_var.get().strip(),
                action_mode="window",
                action_x=self.parse_optional_int(self.app.action_x_var.get()),
                action_y=self.parse_optional_int(self.app.action_y_var.get()),
                ocr_region=RelativeRegion(
                    left=int(self.app.region_left_var.get() or 0),
                    top=int(self.app.region_top_var.get() or 0),
                    width=int(self.app.region_width_var.get() or 0),
                    height=int(self.app.region_height_var.get() or 0),
                ),
                item_points=list(self.app.item_points),
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
        if config.action_delay < 0:
            raise ValueError("取用等待不能小於 0。")
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
        self.app.item_listbox.delete(0, "end")
        for item in self.app.item_points:
            self.app.item_listbox.insert("end", f"{item.name:<8} window=({item.x}, {item.y})")
