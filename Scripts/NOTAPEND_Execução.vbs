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
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectAll
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]").sendVKey 4
session.findById("wnd[4]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[4]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[4]").sendVKey 4
session.findById("wnd[5]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[5]/usr/ctxtDY_PATH").caretPosition = 36
session.findById("wnd[5]").sendVKey 4
session.findById("wnd[6]/usr/ctxtDY_PATH").text = "V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte Albuquerque\Desenvolvimento\Pedro\Relatórios Mensais"
session.findById("wnd[6]/usr/ctxtDY_FILENAME").text = "notapend.XLSX"
session.findById("wnd[6]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[6]/tbar[0]/btn[0]").press
session.findById("wnd[5]/tbar[0]/btn[0]").press
session.findById("wnd[4]").close
session.findById("wnd[3]").close
session.findById("wnd[2]").close
session.findById("wnd[1]").close
session.findById("wnd[1]/tbar[0]/btn[0]").press
