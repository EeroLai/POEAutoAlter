from .config_store import ConfigController
from .runner import AutomationRunner
from .models import AppConfig, AnchorInfo, ItemPoint, RelativeRegion, WindowInfo
from .overlays import PointPicker, RegionPicker
from .services import ClipboardManager, TemplateLocator, WindowManager
from .text_utils import TextNormalizer, normalize_text, parse_target_list
from .win32 import (
    FAST_POLL_INTERVAL,
    HOTKEY_ID_START,
    HOTKEY_ID_STOP,
    MIN_CLIPBOARD_TIMEOUT,
    MOD_NOREPEAT,
    MSG,
    PM_REMOVE,
    USER32,
    VK_F2,
    VK_F3,
    WINDOW_ACTIVATE_DELAY,
    WM_HOTKEY,
    CLIPBOARD_EXTRA_DELAY,
    enable_dpi_awareness,
)
