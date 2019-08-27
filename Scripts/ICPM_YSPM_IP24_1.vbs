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
session.findById("wnd[0]").resizeWorkingPane 95,25,false
session.findById("wnd[0]/tbar[0]/okcd").text = "yspm_ip24"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtSO_WAPOS-LOW").setFocus
session.findById("wnd[0]/usr/ctxtSO_WAPOS-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_SO_WAPOS_%_APP_%-VALU_PUSH").press
session.findById("wnd[0]/usr/ctxtSO_IWERK-LOW").text = "0105"
session.findById("wnd[0]/usr/ctxtSO_IWERK-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtSO_INGPR-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtSO_INGPR-LOW").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtP_LAY").text = "/icpm_bmvm"
session.findById("wnd[0]/usr/ctxtP_LAY").setFocus
session.findById("wnd[0]/usr/ctxtP_LAY").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
