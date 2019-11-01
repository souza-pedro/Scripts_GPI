

'Script para programar os indicadores INOLSI-OK  INNLOE-OK NOTAPEND ORDEMSEMPE BACKLOG IP24

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


'INOLSI

session.findById("wnd[0]/tbar[0]/okcd").text = "iw38"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkDY_OFN").selected = false
session.findById("wnd[0]/usr/ctxtAUART-LOW").text = "zm*"
session.findById("wnd[0]/usr/ctxtAUART-LOW").setFocus
session.findById("wnd[0]/usr/ctxtAUART-LOW").caretPosition = 3
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = "0105"
session.findById("wnd[0]/usr/ctxtIWERK-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtINGPR-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtINGPR-LOW").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btn%_GEWRK_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "020*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "030*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "080*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtVARIANT").text = "/INOLSIBMVM"
session.findById("wnd[0]/usr/ctxtVARIANT").setFocus
session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 11

'Programação INOLSI

session.findById("wnd[0]/usr/ctxtVARIANT").text = "/INOLSIBMVM"
session.findById("wnd[0]/mbar/menu[0]/menu[2]").select
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
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").text = "INOLSI"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").caretPosition = 6
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").text = "INOLSI MENSAL"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").caretPosition = 13
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[13]").press
session.findById("wnd[1]/usr/btnSOFORT_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press 

'Voltar à pagina Inicial

session.findById("wnd[0]/tbar[0]/btn[3]").press

'NOTAPEND

session.findById("wnd[0]/tbar[0]/okcd").text = "ys_notapend"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtSC_IWERK-LOW").text = "0105"
session.findById("wnd[0]/usr/ctxtSC_IWERK-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtSC_INGRP-LOW").text = "0*"
session.findById("wnd[0]/usr/ctxtSC_INGRP-LOW").caretPosition = 2
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtSC_QMART-LOW").text = "z*"
session.findById("wnd[0]/usr/ctxtSC_QMART-LOW").caretPosition = 2
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtSC_ARBPL-LOW").setFocus
session.findById("wnd[0]/usr/ctxtSC_ARBPL-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_SC_ARBPL_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "020*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "030*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "080*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtP_VARI").text = "/MDENG"
session.findById("wnd[0]/usr/ctxtP_VARI").setFocus
session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 6


'NOTAPEND Programação

session.findById("wnd[0]/usr/ctxtP_VARI").text = "/MDENG"
session.findById("wnd[0]/mbar/menu[0]/menu[2]").select
session.findById("wnd[1]/usr/txtPRIPAR_DYN-MAIL").text = "bmvm@petrobras.com.br"
session.findById("wnd[1]/usr/txtPRIPAR_DYN-MAIL").setFocus
session.findById("wnd[1]/usr/txtPRIPAR_DYN-MAIL").caretPosition = 21
session.findById("wnd[1]/tbar[0]/btn[6]").press
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").expandNode "SPOOLREQUEST"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem "PLIST","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem "PLIST","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").topNode = "TEMSE"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem "PLIST","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").text = "NOTAPEND"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").caretPosition = 8
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").text = "NOTAPEND MENSAL"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").caretPosition = 15
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[13]").press
session.findById("wnd[1]/usr/btnSOFORT_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press

'Voltar à pagina Inicial

session.findById("wnd[0]/tbar[0]/btn[3]").press

'ORDEMSEMPE

session.findById("wnd[0]/tbar[0]/okcd").text = "iw38"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkDY_OFN").selected = false
session.findById("wnd[0]/usr/chkDY_MAB").selected = true
session.findById("wnd[0]/usr/ctxtDATUV").text = "01012016"
session.findById("wnd[0]/usr/ctxtSTAI1-LOW").setFocus
session.findById("wnd[0]/usr/ctxtSTAI1-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_STAI1_%_APP_%-VALU_PUSH").press
session.findById("wnd[0]/usr/ctxtAUART-LOW").text = "zm*"
session.findById("wnd[0]/usr/ctxtAUART-LOW").caretPosition = 3
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = "0105"
session.findById("wnd[0]/usr/ctxtIWERK-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtINGPR-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtINGPR-LOW").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "CONF"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "ENTE"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "ENCE"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtVARIANT").text = "/INOLSIBMVM"
session.findById("wnd[0]/usr/ctxtVARIANT").setFocus
session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 11

'ORDEMSEMPE Background

session.findById("wnd[0]/usr/ctxtVARIANT").text = "/INOLSIBMVM"
session.findById("wnd[0]/mbar/menu[0]/menu[2]").select
session.findById("wnd[1]/usr/txtPRIPAR_DYN-MAIL").text = "bmvm@petrobras.com.br"
session.findById("wnd[1]/usr/txtPRIPAR_DYN-MAIL").setFocus
session.findById("wnd[1]/usr/txtPRIPAR_DYN-MAIL").caretPosition = 21
session.findById("wnd[1]/tbar[0]/btn[6]").press
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").expandNode "SPOOLREQUEST"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem "PLIST","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem "PLIST","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").topNode = "TEMSE"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem "PLIST","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").text = "ORDEMSEMPE"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").caretPosition = 10
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").text = "ORDEMSEMPE MENSAL"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").caretPosition = 17
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[13]").press
session.findById("wnd[1]/usr/btnSOFORT_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press

'Voltar à pagina Inicial

session.findById("wnd[0]/tbar[0]/btn[3]").press

'IARI

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

'IARI Background

session.findById("wnd[0]/mbar/menu[0]/menu[2]").select
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
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").text = "IARI"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").caretPosition = 4
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").text = "IARI MENSAL"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").caretPosition = 11
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[13]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/btnSOFORT_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press

'Voltar à pagina Inicial

session.findById("wnd[0]/tbar[0]/btn[3]").press


'BACKLOG

session.findById("wnd[0]/tbar[0]/okcd").text = "iw37n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").text = "zm*"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-HIGH").text = ""
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").caretPosition = 3
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/btn%_S_GEWRK_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "020*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "030*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "080*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AWERK-LOW").text = "0105"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-HIGH").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-HIGH").caretPosition = 0
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").select
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_IWERK-LOW").text = "0105"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_IWERK-LOW").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_IWERK-LOW").caretPosition = 4
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").select
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = "/blog_mdeng"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").caretPosition = 11
session.findById("wnd[0]").sendVKey 0

'Backlog Background

session.findById("wnd[0]/mbar/menu[0]/menu[2]").select
session.findById("wnd[1]/usr/txtPRIPAR_DYN-MAIL").text = "bmvm@petrobras.com.br"
session.findById("wnd[1]/usr/txtPRIPAR_DYN-MAIL").setFocus
session.findById("wnd[1]/usr/txtPRIPAR_DYN-MAIL").caretPosition = 21
session.findById("wnd[1]/tbar[0]/btn[6]").press
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").expandNode "SPOOLREQUEST"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem "PLIST","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem "PLIST","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").topNode = "TEMSE"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem "PLIST","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").text = "Backlog"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").caretPosition = 7
session.findById("wnd[2]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[6]").press
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").expandNode "SPOOLREQUEST"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").topNode = "TEMSE"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").text = "BACKLOG MENSAL"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").caretPosition = 14
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[13]").press
session.findById("wnd[1]/usr/btnSOFORT_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press

'Voltar à pagina Inicial

session.findById("wnd[0]/tbar[0]/btn[3]").press

'IP24 - Seleção dos itens
session.findById("wnd[0]/tbar[0]/okcd").text = "ip24"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtAUART-LOW").setFocus
session.findById("wnd[0]/usr/ctxtAUART-LOW").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]").close
session.findById("wnd[0]/usr/btn%_AUART_%_APP_%-VALU_PUSH").press
session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = "0105"
session.findById("wnd[0]/usr/ctxtIWERK-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtWPGRP-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtWPGRP-LOW").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtAUART-LOW").text = "zm01"
session.findById("wnd[0]/usr/ctxtAUART-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "zm02"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 0
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[2]").close
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "zm03"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "zm04"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "zm05"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").caretPosition = 0
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[2]").close
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "zm06"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "zm07"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").text = "zm08"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btn%_ILART_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "z02"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").caretPosition = 2
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "z03"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "z18"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "z27"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").caretPosition = 0
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[2]").close
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "z28"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "z13"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "z31"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").text = "z32"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").caretPosition = 3
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 1
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").text = "z30"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").caretPosition = 3
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtVARIANT").text = "/MD_ENG"


'IP24 Background

session.findById("wnd[0]/mbar/menu[0]/menu[2]").select
session.findById("wnd[1]/usr/txtPRIPAR_DYN-MAIL").text = "bmvm@petrobras.com.br"
session.findById("wnd[1]/usr/txtPRIPAR_DYN-MAIL").setFocus
session.findById("wnd[1]/usr/txtPRIPAR_DYN-MAIL").caretPosition = 21
session.findById("wnd[1]/tbar[0]/btn[6]").press
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").expandNode "SPOOLREQUEST"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem "PLIST","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem "PLIST","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").topNode = "TEMSE"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem "PLIST","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").text = "IP24"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").caretPosition = 4
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem "PRTXT","Column1"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").text = "IP24 MENSAL"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").caretPosition = 11
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[13]").press
session.findById("wnd[1]/usr/btnSOFORT_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press

'Fechar e abrir Jobs Proprios

session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/mbar/menu[4]/menu[9]").select