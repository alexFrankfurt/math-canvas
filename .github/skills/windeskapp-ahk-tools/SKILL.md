---
name: windeskapp-ahk-tools
description: Three AutoHotkey v2 scripts to focus windeskapp, take a screenshot, or type a fraction at the end of existing text. Use when asked to interact with the running windeskapp application.
---

## Description

Three standalone AHK v2 scripts in `ahk_tools/` that each do one thing. Run them from PowerShell. They print `OK: ...` on success or `ERROR: ...` on failure to stdout.

**AHK runtime**: `C:\Users\alex_\AppData\Local\Programs\AutoHotkey\v2\AutoHotkey64.exe`

## 1. Focus Window

Activates windeskapp and prints its position/size.

```powershell
& "C:\Users\alex_\AppData\Local\Programs\AutoHotkey\v2\AutoHotkey64.exe" "D:\Dev\windeskapp\ahk_tools\focus_window.ahk"
Get-Content "D:\Dev\windeskapp\ahk_tools\focus_window_result.txt"
```

Output: `OK: windeskapp focused at 130,130 size 800x600`

## 2. Take Screenshot

Captures windeskapp to a PNG file. Optional arg: output path (default: `frames\screenshot.png`).

```powershell
# Default path
& "C:\Users\alex_\AppData\Local\Programs\AutoHotkey\v2\AutoHotkey64.exe" "D:\Dev\windeskapp\ahk_tools\take_screenshot.ahk"
Get-Content "D:\Dev\windeskapp\ahk_tools\take_screenshot_result.txt"

# Custom path
& "C:\Users\alex_\AppData\Local\Programs\AutoHotkey\v2\AutoHotkey64.exe" "D:\Dev\windeskapp\ahk_tools\take_screenshot.ahk" "D:\Dev\windeskapp\frames\my_shot.png"
Get-Content "D:\Dev\windeskapp\ahk_tools\take_screenshot_result.txt"
```

Output: `OK: D:\Dev\windeskapp\frames\screenshot.png (15234 bytes)`

## 3. Type Fraction

Moves cursor to end of text and types a fraction (numerator/denominator + Enter to confirm). Does NOT clear existing text.

```powershell
& "C:\Users\alex_\AppData\Local\Programs\AutoHotkey\v2\AutoHotkey64.exe" "D:\Dev\windeskapp\ahk_tools\type_fraction.ahk" "3" "4"
Get-Content "D:\Dev\windeskapp\ahk_tools\type_fraction_result.txt"
```

Args: `<numerator> <denominator>`
Output: `OK: typed fraction 3/4`
