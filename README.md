# POEAutoAlter

POEAutoAlter is a Windows desktop automation tool for Path of Exile.
It can scan item positions, detect text from either OCR or clipboard copy,
and stop automatically when any target keyword is matched.

## Features

- Lock onto a game window by title
- All coordinates are relative to the Path of Exile window
- Multiple item points per scan cycle
- Two detection modes:
  - OCR
  - Clipboard Ctrl+C
- Multiple target keywords supported
- Traditional/Simplified Chinese matching support
- Optional right-click action point
- Optional hold-Shift-through-loop mode for repeated item clicks
- Global hotkeys:
  - F2 to stop
  - F3 to start
- PyAutoGUI failsafe by moving the mouse to the top-left corner
- Local settings are saved to `config.json`

## Requirements

- Windows 10 or Windows 11
- Python 3.11+
- Path of Exile

Recommended:

- Run the game in windowed or borderless mode
- If Path of Exile is started as Administrator, run this tool with the same privilege level

## Install

```powershell
python -m pip install -r requirements.txt
```

## Run

```powershell
python app.py
```

or:

```powershell
run.bat
```

## Quick Start

1. Open Path of Exile.
2. Launch this tool.
3. Keep `Window Title Keyword` as `Path of Exile`, or use `Grab Foreground Window`.
4. Click `Test Window` and confirm the correct game window is found.
5. Choose a detection mode:
   - `OCR`
   - `Clipboard Ctrl+C`
6. Enter one or more target keywords.
   - Supported separators: comma, semicolon, or `|`
   - Matching any keyword will stop the automation
7. Optionally set `Right Click X/Y`.
   - If left empty, the tool only performs hover + detection
8. Optionally enable `Hold Shift Through Loop`.
   - After the first right-click pickup, the tool keeps Shift held during the loop and uses left-click on items
   - In clipboard mode, Shift remains held even while sending `Ctrl+C`
9. Add one or more item points.
10. Adjust timing settings if needed:
   - `hover delay`
   - `click delay`
   - `click jitter`
   - `cycle delay`
11. Click `Start`.

## Detection Modes

### OCR mode

Use this when the item text appears in a fixed region of the game window.

Setup:

1. Click `Pick OCR Region`
2. Drag over the text area inside the game window
3. Click `Test Detection`

Notes:

- `OCR Left / Top / Width / Height` are relative to the game window
- Smaller OCR regions are usually more reliable

### Clipboard Ctrl+C mode

Use this when hovering an item and pressing `Ctrl+C` copies item text to the clipboard.

Flow:

1. Move to an item point
2. Send `Ctrl+C`
3. Read clipboard text
4. Match against the target keyword list

Notes:

- This mode does not depend on the OCR region
- If clipboard reads are unstable, increase `hover delay` or `click delay`

## Automation Flow

Each cycle works like this:

1. Resolve the Path of Exile window
2. Move to an item point
3. Run a detection check before clicking
4. If a keyword is matched, stop immediately
5. If not matched:
   - Normal mode: if a right-click point is set, right-click that point and left-click the item
   - Hold-Shift mode: right-click the action point once, then keep Shift held and continue left-clicking items
   - If no right-click point is set, skip the click step
6. Run detection again
7. Continue to the next item or next cycle

## Coordinate System

All coordinates are relative to the top-left corner of the Path of Exile window:

- Right-click point
- Item points
- OCR region

This keeps the setup stable even when the game window moves.

## Hotkeys and Safety

- `F2`: stop
- `F3`: start
- Mouse to top-left corner: trigger PyAutoGUI failsafe

## Local Config

The app stores local runtime settings in `config.json`, including:

- window title
- detection mode
- target keywords
- right-click point
- OCR region
- item points
- timing and click jitter

`config.json` is intentionally ignored by Git.

## Troubleshooting

### OCR does not detect text

- Re-pick a smaller OCR region
- Make sure the text really appears in a fixed part of the window
- If Path of Exile supports item copy with `Ctrl+C`, prefer clipboard mode

### F2 or F3 does not respond

- Make sure the tool is still running
- Run the tool with the same privilege level as the game

### Clipboard mode does not copy item text

- Make sure the cursor is really hovering the item
- Increase `hover delay`, for example to `0.1` or `0.2`
- Increase `click delay` slightly so the game has time to update

## Notes

This project is currently tailored to the existing Path of Exile workflow in this repo.
If the workflow changes later, it can be extended further with:

- different click sequences
- more hotkeys
- EXE packaging
- GitHub Actions build automation
