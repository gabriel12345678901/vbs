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









Session.findById("wnd[0]").maximize

'login lucineia'
'session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "tr503769"
'session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "equipe03"

'login junior'
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "tr498602"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "jno058pr"

session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus
session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 8
session.findById("wnd[0]").sendVKey 0


'Session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
Session.findById("wnd[0]/tbar[0]/okcd").Text = "mb52"
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/usr/ctxtWERKS-LOW").SetFocus
'clica no botão para inserir diversos centros
Session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press
'insere os centros
 Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "tlpr"
'Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "TLSC"
'Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "tlrs"

'clica no botão executar f8
Session.findById("wnd[1]/tbar[0]/btn[8]").press
'preenche campo deposito
'saldo volante
 session.findById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").press
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "DT01"
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "DT02"
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "DT03"
 'estoque
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "CA01"
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "CA02"
 session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "CA03"


'clica no botão executar f8
Session.findById("wnd[0]/tbar[1]/btn[8]").press
Session.findById("wnd[0]/tbar[1]/btn[8]").press
'salva relatório na máquina
'Session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").Select

    
Set SapGuiAuto = Nothing
Set MyApp = Nothing
Set MyConnection = Nothing
Set Session = Nothing