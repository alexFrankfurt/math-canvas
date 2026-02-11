; Type a fraction at the end of existing text in windeskapp
; Usage: AutoHotkey64.exe type_fraction.ahk <numerator> <denominator>
; Example: AutoHotkey64.exe type_fraction.ahk 3 4
#Requires AutoHotkey v2.0

resultFile := A_ScriptDir "\type_fraction_result.txt"

WriteResult(msg) {
    global resultFile
    try FileDelete(resultFile)
    FileAppend(msg, resultFile)
}

if A_Args.Length < 2 {
    WriteResult("ERROR: usage: type_fraction.ahk <numerator> <denominator>")
    ExitApp 1
}

numerator := A_Args[1]
denominator := A_Args[2]

if !WinExist("ahk_exe windeskapp.exe") {
    WriteResult("ERROR: windeskapp not found")
    ExitApp 1
}

WinActivate("ahk_exe windeskapp.exe")
if !WinWaitActive("ahk_exe windeskapp.exe",, 5) {
    WriteResult("ERROR: could not activate windeskapp")
    ExitApp 1
}

; Focus the RichEdit control
try {
    editHwnd := ControlGetHwnd("RICHEDIT50W1", "ahk_exe windeskapp.exe")
    ControlFocus(editHwnd, "ahk_exe windeskapp.exe")
} catch {
    WriteResult("ERROR: RichEdit control not found")
    ExitApp 1
}
Sleep(200)

; Move cursor to end of text (Ctrl+End)
Send("^{End}")
Sleep(200)

; Type numerator, then / to trigger fraction, then denominator, then Enter to confirm
SendText(numerator)
Sleep(200)
SendText("/")
Sleep(400)
SendText(denominator)
Sleep(300)
Send("{Enter}")
Sleep(300)

WriteResult("OK: typed fraction " numerator "/" denominator)
ExitApp 0
