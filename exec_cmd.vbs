const DontWaitUntilFinished = false, ShowWindow = 1, DontShowWindow = 0, WaitUntilFinished = true
set oShell = WScript.CreateObject("WScript.Shell")
command = "cmd /c C:\windows\system32\wscript.exe C:\SCRIPT.vbs " & args
oShell.Run command, DontShowWindow, WaitUntilFinished