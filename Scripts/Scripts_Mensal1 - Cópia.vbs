

'Script para programar os indicadores INOLSI INNLOE NOTAPEND ORDEMSEMPE BACKLOG IP24

'Cabeçalho. Inciar com uma única janela aberta no SAP

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

'INNLOE

session.findById("wnd[0]/tbar[0]/okcd").text = "iw28"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkDY_OFN").selected = false
session.findById("wnd[0]/usr/ctxtQMART-LOW").text = "z*"
session.findById("wnd[0]/usr/ctxtQMART-LOW").setFocus
session.findById("wnd[0]/usr/ctxtQMART-LOW").caretPosition = 2
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = "0105"
session.findById("wnd[0]/usr/ctxtIWERK-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtARBPL-LOW").setFocus
session.findById("wnd[0]/usr/ctxtARBPL-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_ARBPL_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "020*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "030*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "080*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtVARIANT").text = "/mdeng"
session.findById("wnd[0]/usr/ctxtVARIANT").setFocus
session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 6
session.findById("wnd[0]").sendVKey 0

'Programação em BackGround INNLOE

session.findById("wnd[0]/mbar/menu[0]/menu[2]").select
session.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").text = "mail"
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/usr/txtPRIPAR_DYN-MAIL").text = "bmvm@petrobras.com.br"
session.findById("wnd[1]/usr/txtPRIPAR_DYN-MAIL").setFocus
session.findById("wnd[1]/usr/txtPRIPAR_DYN-MAIL").caretPosition = 21
session.findById("wnd[1]/tbar[0]/btn[6]").press
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").expandNode "OUTPUT"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").expandNode "SPOOLREQUEST"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem "PLIST","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem "PLIST","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").topNode = "OUTPUT"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem "PLIST","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").text = "INNLOE"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").caretPosition = 6
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").text = "INNLOE Mensal"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").caretPosition = 13
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[13]").press
session.findById("wnd[1]/usr/btnSOFORT_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press

'Voltar à pagina Inicial

session.findById("wnd[0]/tbar[0]/btn[3]").press




