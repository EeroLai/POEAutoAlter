from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class ItemPoint:
    name: str
    x: int
    y: int

    @classmethod
    def from_dict(cls, data: dict) -> "ItemPoint":
        return cls(
            name=str(data.get("name", "??")),
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
    action_delay: float = 0.12
    click_jitter: int = 2
    human_delay: bool = True
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
            action_delay=float(data.get("action_delay", 0.12)),
            click_jitter=int(data.get("click_jitter", 2)),
            human_delay=bool(data.get("human_delay", True)),
            hold_shift_loop=bool(data.get("hold_shift_loop", False)),
            cycle_delay=float(data.get("cycle_delay", 0.8)),
            target_text=str(data.get("target_text", "")),
            action_mode=str(data.get("action_mode", "window")),
            action_x=data.get("action_x"),
            action_y=data.get("action_y"),
            ocr_region=RelativeRegion.from_dict(data.get("ocr_region")),
            item_points=[ItemPoint.from_dict(item) for item in data.get("item_points", [])],
        )
