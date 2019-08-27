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
session.findById("wnd[0]/tbar[0]/okcd").text = "iw66"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtQMART-LOW").text = "zr"
session.findById("wnd[0]/usr/ctxtQMART-LOW").setFocus
session.findById("wnd[0]/usr/ctxtQMART-LOW").caretPosition = 2
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = "0105"
session.findById("wnd[0]/usr/ctxtIWERK-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkDY_QMSM").selected = false
session.findById("wnd[0]/usr/ctxtVARIANT").text = "/MDENG"
session.findById("wnd[0]/usr/ctxtVARIANT").setFocus
session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 6
session.findById("wnd[0]").sendVKey 0
