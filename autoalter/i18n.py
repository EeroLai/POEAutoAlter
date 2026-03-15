from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path

LOCALES_DIR = Path(__file__).resolve().parent.parent / 'locales'
SUPPORTED_LANGUAGES = ('zh', 'en')
LOCALE_FILES = {
    'zh': LOCALES_DIR / 'zh-TW.json',
    'en': LOCALES_DIR / 'en.json',
}
ZH_OF = '對 '
FULLWIDTH_COLON = '：'
IDEOGRAPHIC_FULL_STOP = '。'
FULLWIDTH_COMMA = '，'


@lru_cache(maxsize=1)
def _load_locales() -> dict[str, dict]:
    locales: dict[str, dict] = {}
    for language, path in LOCALE_FILES.items():
        locales[language] = json.loads(path.read_text(encoding='utf-8'))
    return locales


LOCALES = _load_locales()
LANGUAGE_NAMES = {
    language: LOCALES[language]['language_name']
    for language in SUPPORTED_LANGUAGES
}
UI_TEXT = {
    language: LOCALES[language]['ui']
    for language in SUPPORTED_LANGUAGES
}


def translate_runtime_text(language: str, text: str) -> str:
    if language != 'en':
        return text

    locale = LOCALES['en']
    runtime_map = locale.get('runtime_map', {})
    if text in runtime_map:
        return runtime_map[text]

    for item in locale.get('runtime_prefix_map', []):
        source = item['source']
        if text.startswith(source):
            return item['target'] + text[len(source):]

    tokens = locale.get('runtime_tokens', {})
    scan_prefix = tokens.get('scan_cycle_prefix', '')
    scan_suffix = tokens.get('scan_cycle_suffix', '')
    if scan_prefix and scan_suffix and text.startswith(scan_prefix) and text.endswith(scan_suffix):
        number = text[len(scan_prefix):-len(scan_suffix)]
        return tokens.get('scan_cycle_format', 'Scanning, cycle {number}').format(number=number)

    cycle_start_prefix = tokens.get('cycle_start_prefix', '')
    cycle_start_split = tokens.get('cycle_start_split', '')
    if cycle_start_prefix and cycle_start_split and text.startswith(cycle_start_prefix) and cycle_start_split in text:
        number, rest = text[len(cycle_start_prefix):].split(cycle_start_split, 1)
        return tokens.get('cycle_start_format', 'Cycle {number} started, window={rest}').format(number=number, rest=rest)

    before_hover = tokens.get('before_hover')
    after_hover = tokens.get('after_hover')
    if before_hover and text.startswith(before_hover):
        return text.replace(before_hover, 'Before hover ', 1)
    if after_hover and text.startswith(after_hover):
        return text.replace(after_hover, 'After hover ', 1)

    before = tokens.get('before')
    after = tokens.get('after')
    clipboard = tokens.get('clipboard')
    clipboard_en = tokens.get('clipboard_en', 'Clipboard')
    no_content = tokens.get('no_content')
    no_content_en = tokens.get('no_content_en', 'No content captured')
    if before and clipboard and text.startswith(before + clipboard + '['):
        return text.replace(before, 'Before ', 1).replace(clipboard, clipboard_en, 1).replace(no_content, no_content_en)
    if after and clipboard and text.startswith(after + clipboard + '['):
        return text.replace(after, 'After ', 1).replace(clipboard, clipboard_en, 1).replace(no_content, no_content_en)

    ocr = tokens.get('ocr')
    no_text = tokens.get('no_text')
    no_text_en = tokens.get('no_text_en', 'No text captured')
    if before and ocr and text.startswith(before + ocr + '['):
        return text.replace(before, 'Before ', 1).replace(no_text, no_text_en)
    if after and ocr and text.startswith(after + ocr + '['):
        return text.replace(after, 'After ', 1).replace(no_text, no_text_en)

    right_click_prefix = tokens.get('right_click_prefix')
    if right_click_prefix and text.startswith(right_click_prefix):
        return text.replace(right_click_prefix, tokens.get('right_click_en', 'Right-clicked action point: '), 1)

    left_click_split = tokens.get('left_click_split')
    if left_click_split and text.startswith(ZH_OF) and left_click_split in text:
        return text.replace(ZH_OF, tokens.get('left_click_en_prefix', 'Left-clicked '), 1).replace(left_click_split, ': ', 1)

    shift_left_click_split = tokens.get('shift_left_click_split')
    if shift_left_click_split and text.startswith(ZH_OF) and shift_left_click_split in text:
        return text.replace(ZH_OF, tokens.get('shift_left_click_en_prefix', 'Shift+left-clicked '), 1).replace(shift_left_click_split, ': ', 1)

    shift_primed_prefix = tokens.get('shift_primed_prefix')
    if shift_primed_prefix and text.startswith(shift_primed_prefix):
        return text.replace(shift_primed_prefix, tokens.get('shift_primed_en', 'Shift loop primed action point: '), 1)

    action_updated_prefix = tokens.get('action_updated_prefix')
    if action_updated_prefix and text.startswith(action_updated_prefix):
        return text.replace(action_updated_prefix, tokens.get('action_updated_en', 'Action point updated: '), 1)

    ocr_region_updated_prefix = tokens.get('ocr_region_updated_prefix')
    if ocr_region_updated_prefix and text.startswith(ocr_region_updated_prefix):
        return text.replace(ocr_region_updated_prefix, tokens.get('ocr_region_updated_en', 'OCR region updated: '), 1)

    added_prefix = tokens.get('added_prefix')
    if added_prefix and text.startswith(added_prefix):
        return text.replace(added_prefix, tokens.get('added_en', 'Added '), 1)

    registered_hotkeys_prefix = tokens.get('registered_hotkeys_prefix')
    if registered_hotkeys_prefix and text.startswith(registered_hotkeys_prefix):
        payload = text.split(FULLWIDTH_COLON, 1)[1].rstrip(IDEOGRAPHIC_FULL_STOP)
        payload = payload.replace(tokens.get('f2_stop', 'F2=停止'), tokens.get('f2_stop_en', 'F2=Stop'))
        payload = payload.replace(tokens.get('f3_start', 'F3=開始'), tokens.get('f3_start_en', 'F3=Start'))
        payload = payload.replace(FULLWIDTH_COMMA, ', ')
        return f"{tokens.get('registered_hotkeys_en', 'Registered global hotkeys: ')}{payload}."

    automation_started_prefix = tokens.get('automation_started_prefix')
    if automation_started_prefix and text.startswith(automation_started_prefix):
        text = text.replace(automation_started_prefix, tokens.get('automation_started_en', 'Automation started, mode='), 1)
        text = text.replace(tokens.get('pickup_delay', '取用等待='), tokens.get('pickup_delay_en', 'pickup_delay='), 1)
        text = text.replace(tokens.get('click_jitter_label', '點擊浮動='), tokens.get('click_jitter_label_en', 'click_jitter='), 1)
        return text

    hotkey_help_runtime = tokens.get('hotkey_help_runtime')
    if hotkey_help_runtime and text.startswith(hotkey_help_runtime):
        return tokens.get('hotkey_help_runtime_en', 'F2 stops globally, F3 starts globally.')

    failsafe = tokens.get('failsafe')
    if failsafe and text.startswith(failsafe):
        return tokens.get('failsafe_en', 'Moved mouse to top-left corner and triggered PyAutoGUI failsafe. Automation stopped.')

    return text
