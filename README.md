# POEAutoAlter

POEAutoAlter is a Windows desktop automation tool for Path of Exile that automates Alteration Orb item rolling by using clipboard item text detection.

## Repo Description

Clipboard-driven Path of Exile Alteration Orb automation tool with window monitoring, multi-item routing, and match-based stopping.

## What It Does

- Monitors the `Path of Exile` window automatically
- Enables `Start` only when the game window is detected
- Stops automatically if the monitored window disappears while running
- Uses `Ctrl+C` item copy instead of OCR
- Matches copied item text against one or more target keywords
- Processes item points in order
- Keeps working on `Item1` until it matches, then moves to `Item2`, then `Item3`, and so on
- Stops when all configured item points are completed
- Supports optional Alteration Orb position clicks
- Supports optional hold-Shift mode for repeated crafting flow
- Adds click jitter and optional humanized delay
- Includes stale clipboard protection to avoid extra clicks when the item text has not updated yet
- Supports Chinese and English UI

## Requirements

- Windows 10 or Windows 11
- Python 3.11+
- Path of Exile

Recommended:

- Run the game in windowed or borderless mode
- Run this tool with the same privilege level as the game

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
2. Launch the tool.
3. Confirm the window status shows that `Path of Exile` is being monitored.
4. Enter one or more target keywords.
   Matching uses the clipboard item text and stops on any matched keyword.
5. Set the Alteration Orb position.
6. Add one or more item points.
7. Adjust timing settings if needed.
8. Click `Start` or press `F3`.

## Detection

This project currently uses clipboard detection only.

Flow:

1. Move the cursor to an item point
2. Send `Ctrl+C`
3. Read the copied item text from the clipboard
4. Match the result against the target keyword list

Supported separators for the target keyword list:

- `,`
- `;`
- `|`

## Automation Flow

For each configured item point:

1. Hover the item
2. Copy item text with `Ctrl+C`
3. Check for a target keyword
4. If matched, mark that item point as complete and move to the next item point
5. If not matched, use the Alteration Orb position and click the item
6. Copy again and re-check
7. If the copied text is still identical to the previous result, wait briefly and re-check instead of immediately adding another click

When all item points are completed, the run stops.

## Window Monitoring

The app checks for the `Path of Exile` window once per second.

Behavior:

- If the window is found, `Start` is enabled
- If the window is not found, `Start` is disabled
- If the window disappears during a run, the automation stops immediately

## Coordinate System

All configured positions are relative to the top-left corner of the monitored `Path of Exile` window:

- Alteration Orb position
- Item points

## Timing Options

- `Humanized delay`: adds a small random extra delay to waits
- `Hover delay`: delay before clipboard reading after hover
- `Click delay`: delay after an item click
- `Pickup delay`: delay after picking the Alteration Orb
- `Click jitter`: random click offset in pixels
- `Cycle delay`: kept for compatibility, but the current flow finishes after all configured item points are processed

## Shift Mode

If `Hold Shift through loop` is enabled:

- The tool primes the Alteration Orb action once
- Keeps Shift held during the item processing flow
- Releases Shift on stop, on match-stop, or on close

## Hotkeys and Safety

- `F2`: stop
- `F3`: start
- Move the mouse to the top-left corner to trigger PyAutoGUI failsafe

## Local Config

Local settings are stored in `config.json`.

Typical values include:

- target keywords
- Alteration Orb position
- item points
- timing settings
- click jitter
- language

`config.json` is ignored by Git.

## Troubleshooting

### The tool clicks too early after the item changes

- Increase `Hover delay`
- Increase `Pickup delay`
- Keep the stale clipboard protection enabled by using the current version

### Clipboard text does not update correctly

- Make sure the cursor is actually hovering the item
- Make sure the game supports `Ctrl+C` item copy in the current state
- Run the tool with the same privilege level as the game

### Start is disabled

- Make sure Path of Exile is open
- Make sure the game window title still contains `Path of Exile`
- Wait for the next 1-second monitor refresh or click `Refresh now`

## Notes

This repository is focused on a practical Path of Exile Alteration Orb workflow rather than a generic game bot framework.
