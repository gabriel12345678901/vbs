'Dim WShell
'Set WShell = CreateObject("WScript.Shell")
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "C:\SAP\FrontEnd\SapGui\saplogon.exe"


WScript.Sleep 7000
WshShell.SendKeys "{enter}"
WScript.Sleep 10000


If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

WScript.Sleep 2000



'Session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "tr503769"
'Session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "equipe03"

Session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "tr498602"
Session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "jno058pr"

'enter'
Session.findById("wnd[0]").sendVKey 0

WScript.echo "Login realizado."