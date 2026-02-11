; Focus windeskapp window
; Usage: AutoHotkey64.exe focus_window.ahk
#Requires AutoHotkey v2.0

resultFile := A_ScriptDir "\focus_window_result.txt"

WriteResult(msg) {
    global resultFile
    try FileDelete(resultFile)
    FileAppend(msg, resultFile)
}

if !WinExist("ahk_exe windeskapp.exe") {
    WriteResult("ERROR: windeskapp not found")
    ExitApp 1
}

WinActivate("ahk_exe windeskapp.exe")
if !WinWaitActive("ahk_exe windeskapp.exe",, 5) {
    WriteResult("ERROR: could not activate windeskapp")
    ExitApp 1
}

WinGetPos(&x, &y, &w, &h, "ahk_exe windeskapp.exe")
WriteResult("OK: windeskapp focused at " x "," y " size " w "x" h)
ExitApp 0
