; Take screenshot of windeskapp window
; Usage: AutoHotkey64.exe take_screenshot.ahk [output_path]
; Default output: frames\screenshot.png
#Requires AutoHotkey v2.0

resultFile := A_ScriptDir "\take_screenshot_result.txt"

WriteResult(msg) {
    global resultFile
    try FileDelete(resultFile)
    FileAppend(msg, resultFile)
}

outPath := A_Args.Length > 0 ? A_Args[1] : A_ScriptDir "\..\frames\screenshot.png"

if !WinExist("ahk_exe windeskapp.exe") {
    WriteResult("ERROR: windeskapp not found")
    ExitApp 1
}

WinActivate("ahk_exe windeskapp.exe")
WinWaitActive("ahk_exe windeskapp.exe",, 3)
WinGetPos(&wx, &wy, &ww, &wh, "ahk_exe windeskapp.exe")

; ---- GDI+ helpers ----
DllCall("LoadLibrary", "Str", "gdiplus")
si := Buffer(A_PtrSize = 8 ? 24 : 16, 0)
NumPut("UInt", 1, si, 0)
pToken := 0
DllCall("gdiplus\GdiplusStartup", "Ptr*", &pToken, "Ptr", si, "Ptr", 0)

; Capture screen region
hScreen := DllCall("GetDC", "Ptr", 0, "Ptr")
hMemDC := DllCall("CreateCompatibleDC", "Ptr", hScreen, "Ptr")
hBmp := DllCall("CreateCompatibleBitmap", "Ptr", hScreen, "Int", ww, "Int", wh, "Ptr")
DllCall("SelectObject", "Ptr", hMemDC, "Ptr", hBmp)
DllCall("BitBlt", "Ptr", hMemDC, "Int", 0, "Int", 0, "Int", ww, "Int", wh, "Ptr", hScreen, "Int", wx, "Int", wy, "UInt", 0x00CC0020)
DllCall("ReleaseDC", "Ptr", 0, "Ptr", hScreen)

; Convert to GDI+ bitmap and save as PNG
pBitmap := 0
DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "Ptr", hBmp, "Ptr", 0, "Ptr*", &pBitmap)
DllCall("DeleteObject", "Ptr", hBmp)
DllCall("DeleteDC", "Ptr", hMemDC)

; Find PNG encoder
; ImageCodecInfo: 2 GUIDs (32) + 5 ptrs + 4 DWORDs (16) + 2 ptrs = 48 + 7*PtrSize
count := 0, size := 0
DllCall("gdiplus\GdipGetImageEncodersSize", "UInt*", &count, "UInt*", &size)
ci := Buffer(size)
DllCall("gdiplus\GdipGetImageEncoders", "UInt", count, "UInt", size, "Ptr", ci)
stride := 48 + 7 * A_PtrSize
pEncoder := 0
Loop count {
    offset := (A_Index - 1) * stride
    pMime := NumGet(ci, offset + 32 + A_PtrSize * 4, "UPtr")
    if !pMime
        continue
    mime := StrGet(pMime, "UTF-16")
    if (mime = "image/png") {
        pEncoder := ci.Ptr + offset
        break
    }
}

; Ensure output directory exists
SplitPath(outPath,, &dir)
if dir && !DirExist(dir)
    DirCreate(dir)

DllCall("gdiplus\GdipSaveImageToFile", "Ptr", pBitmap, "WStr", outPath, "Ptr", pEncoder, "Ptr", 0)
DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
DllCall("gdiplus\GdiplusShutdown", "Ptr", pToken)

if FileExist(outPath) {
    fObj := FileOpen(outPath, "r")
    fSize := fObj.Length
    fObj.Close()
    WriteResult("OK: " outPath " (" fSize " bytes)")
    ExitApp 0
} else {
    WriteResult("ERROR: screenshot file not created")
    ExitApp 1
}
