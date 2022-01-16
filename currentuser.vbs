On Error Resume Next

Dim objShell, strTemp
Set objShell = WScript.CreateObject("WScript.Shell")

strTemp = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
WScript.Echo "test 1: " & objShell.RegRead(strTemp) 

strTemp = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Persistence"
WScript.Echo "test 2: " & objShell.RegRead(strTemp) 