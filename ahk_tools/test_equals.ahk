#Requires AutoHotkey v2.0
if !WinWaitActive("ahk_exe windeskapp.exe",, 5) {
    ExitApp 1
}
editHwnd := ControlGetHwnd("RICHEDIT50W1", "ahk_exe windeskapp.exe")
ControlFocus(editHwnd, "ahk_exe windeskapp.exe")
Sleep(200)
Send("^{End}")
Sleep(200)
Send("{Enter}")
Sleep(200)
SendText("10")
Sleep(200)
SendText("/")
Sleep(400)
SendText("4")
Sleep(300)
Send("=")
Sleep(500)
ExitApp 0
