set WshShell = Wscript.CreateObject("WScript.Shell")
cmdID = WshShell.run("Anarede.exe")
WshShell.AppActivate(cmdID)
wscript.sleep 1000
WshShell.SendKeys "{ENTER}"
wscript.sleep 500
WshShell.SendKeys "{ENTER}"
wscript.sleep 200
WshShell.SendKeys "%"
wscript.sleep 200
WshShell.SendKeys "c"
wscript.sleep 200
WshShell.SendKeys "c"
wscript.sleep 1000
WshShell.SendKeys "comando.dat"
wscript.sleep 200
WshShell.SendKeys "{ENTER}"
wscript.sleep 5000
