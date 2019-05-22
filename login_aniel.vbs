Dim WShell
Set WShell = CreateObject("WScript.Shell")


WShell.Run "C:\sinapse\AnielT5\ANIELTV.exe"
WScript.Sleep 7000
WShell.SendKeys "^l"
WShell.SendKeys "jesus123"
WScript.Sleep 500
WShell.SendKeys "{enter}"
WScript.Sleep 500
WShell.SendKeys "{enter}"

WScript.Sleep 500








