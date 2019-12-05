
# Rotina P/ Verificar ordens ainda em aberto no SAP da Programação semanal
#
# Pede a semana para análise, vai na rede e abre o arquivo oficial correto. Depois extrai as confirmações e
# busca quais ainda estão em aberto no SAP.

import os
import pandas as pd
import win32com.client
import fnmatch
import glob
import easygui as g



def carrega_iw37():
    import sys
    import win32com.client
    import datetime

    try:

        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return

        connection = application.Children(0)
        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return

        #>Insert your SAP GUI Scripting code here<
        session.findById("wnd[0]/tbar[0]/okcd").text = "iw37n"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/chkSP_OFN").selected = 0
        session.findById("wnd[0]/usr/chkSP_IAR").selected = -1

        #session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press()
        #session.findById("wnd[1]/tbar[0]/btn[24]").press()
        #session.findById("wnd[1]/tbar[0]/btn[8]").press()

        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").select()
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_IWERK-LOW").text = "0105"
        session.findById("wnd[0]").sendVKey(0)
        #Aba Operação
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").select()
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAEX-LOW").text = "CONF"
        #Aba Outros
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").select()
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = r"/MD_progsem"
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        #Filtra CONF e ELIM
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "STTXT"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleColumn = "ARBEI"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn("STTXT")
        session.findById("wnd[0]/tbar[1]/btn[38]").press()
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").select()
        session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "*CONF*"
        session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").caretPosition = 6
        session.findById("wnd[2]/tbar[0]/btn[8]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(-1, "V_STTXT")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn("V_STTXT")
        session.findById("wnd[0]/tbar[1]/btn[38]").press()
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").select()
        session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "*ELIM*"
        session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").caretPosition = 6
        session.findById("wnd[2]/tbar[0]/btn[8]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        #Salva Arquivo
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(1, "ISMNW")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte Albuquerque\Desenv\Pedro\Pycharm\Confirmação Semanal"
        n_arquivo = "Conf_Semanal_" + datetime.datetime.today().strftime("%y-%m-%d") + ".xlsx"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = n_arquivo
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[11]").press()






    except:
        print(sys.exc_info()[0])

    finally:
        return n_arquivo
        session = None
        connection = None
        application = None
        SapGuiAuto = None




#Caminho das pastas padrão
p_salvamento = r"V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte Albuquerque\Desenv\Pedro\Pycharm\Confirmação Semanal"
c_programacao = r"V:\COMPARTILHADO_CSC-SSE_NSIF\NP-1\GPI\1 - Programações Semanais\INDUSTRIAL\2019"
L = []
c_sem = []

#with os.scandir(c_programação) as entries:
#    for entry in entries:
#        print(entry.name)
#        L.append(entry.name)


#Insere a semana da análise e busca na pasta IND
sem = input("digite a semana para análise")

sem_string = c_programacao + r"\*" + sem + r"*"

c_sem = glob.glob(sem_string)
c_sem = c_sem[0]

#Seleciona arquivo com string "original"
for filename in os.listdir(c_sem):
    if fnmatch.fnmatch(filename, "*oficial*.xlsm"):
        L.append(filename)
        print(filename)

#Trata quando acha mais de um arquivo
if len(L) != 1:
    msg = "Mais de um arquivo foi encontrado. Selecione o arquivo desejado:"
    title = "Escolha um arquivo para carga"
    choice = g.choicebox(msg, title, L)
    L = choice
if type(L) == list:
    L = L[0]



#Abrir as ordens no pandas e copiar na área de transferência
c_arquivo = os.path.join(c_sem, L)
dados = pd.read_excel(c_arquivo, skiprows=5, sheet_name='Programação Semanal', convert_float=False, skip_blank_lines=True)
dados.dropna(subset=['CONFIRMAÇÃO'], inplace=True)
dados1 = dados.astype({"CONFIRMAÇÃO": "int32"})
dados1["CONFIRMAÇÃO"].to_clipboard(excel=True, index=False, header=None)

#Abrir SAp, IW37N, pegar só as ordens que estão aidna em aberto
c_resultado = carrega_iw37()
print(c_resultado)








