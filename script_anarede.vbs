set WshShell = Wscript.CreateObject("WScript.Shell")
cmdID = WshShell.run("[PROGRAMA_DESEJADO].exe")
WshShell.AppActivate(cmdID)
'Seguindo um passo para abrir um arquivo dentro do programa que quero executar
wscript.sleep 1000
WshShell.SendKeys "{ENTER}"
wscript.sleep 500
WshShell.SendKeys "{ENTER}"
wscript.sleep 200
WshShell.SendKeys "%" 'botão Alt
wscript.sleep 200
WshShell.SendKeys "c"
wscript.sleep 200
WshShell.SendKeys "c"
wscript.sleep 1000
WshShell.SendKeys "[NOME_ARQUIVO].[EXTENSÃO]"
wscript.sleep 200
WshShell.SendKeys "{ENTER}"
wscript.sleep 5000
