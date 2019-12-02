import sys
import easygui as g
import os
import pandas as pd

# Escolha da pasta de Destino

def main():

    #Escolhe a pasta padrão
    escolher_pasta(c_origem, c_destino)

    #Extrai Nº das Ordens
    Lista_Ordens(c_origem)

    #Anexa arquivos
    anexa_SAP(c_origem)

    #Renomeia arquivos e coloca em pasta destino
    renomeia(c_origem, c_destino)


main()

















def escolher_pasta(c_origem, c_destino):

    c_pasta_origem = r"V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte " \
             r"Albuquerque\Desenvolvimento\Pedro\Pycharm\Anexar_Ordem_SAP\A Transferir"


    while 1:

        msg = "Escolha a Pasta de Origem dos arquivos. Gostaria de Selecionar a pasta ou usar a pasta padrão?" \
             r"      Pasta Padrão: " + c_pasta_origem
        title = "Escolha Pasta Origem"
        choices = ["Usar Padrão", "Escolher"]
        choice = g.choicebox(msg, title, choices)

        # note that we convert choice to string, in case
        # the user cancelled the choice, and we got None.

        if  choice == "Escolher":
                choice = g.diropenbox()
                c_pasta_origem = choice

        #g.msgbox("Você escolheu: " + str(c_pasta_origem), "Resultado Escolha")

        msg = "Gostaria de Continuar?   Você escolheu: " + str(c_pasta_origem)
        title = "Please Confirm"
        if g.ccbox(msg, title):     # show a Continue/Cancel dialog
            pass  # user chose Continue
            break
        else:
            choice = "" # user chose cancel
    # user chose Cancel


    #Escolha da pasta de Saída


    c_pasta_destino = r"V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte " \
             r"Albuquerque\Desenvolvimento\Pedro\Pycharm\Anexar_Ordem_SAP\OK"


    while 1:

        msg = "Escolha a Pasta de Destino. Gostaria de Selecionar a pasta ou usar a pasta padrão?" \
             r"      Pasta Padrão: " + c_pasta_destino
        title = "Escolha Pasta Destino"
        choices = ["Usar Padrão", "Escolher"]
        choice = g.choicebox(msg, title, choices)

        # note that we convert choice to string, in case
        # the user cancelled the choice, and we got None.

        if  choice == "Escolher":
                choice = g.diropenbox()
                c_pasta_destino = choice

        #g.msgbox("Você escolheu: " + str(c_pasta_destino), "Resultado Escolha")

        msg = "Gostaria de Continuar?   Você escolheu: " + str(c_pasta_destino)
        title = "Please Confirm"
        if g.ccbox(msg, title):     # show a Continue/Cancel dialog
            pass  # user chose Continue
            break
        else:
            choice = "" # user chose cancel


    return c_pasta_origem, c_pasta_destino



escolher_pasta()


def Lista_Ordens(c_origem)

#Transformando Nomes dos aquivos em lista de Ordens.
#Nome do arquivo deve estar no formato XXXXXXXXXX_xx-xx-xx_OPER-XX_C (Nº Ordem_dia-mes_ano_OPER-Nº Oper_C)

a = os.listdir(c_pasta_origem)

c = list(range(len(a)))

for f in range(0, len(a)):
    #low = str.find(a[f], " - ") + 3
    low = 0
    upper = str.find(a[f], "_")
    b = a[f]
    c[f] = b[low:upper]
    #print(c, low, upper)

#Ordens = c
ordens = pd.DataFrame(c)
ordens.to_clipboard(index=False, header=False)




print("FIM ")
print("Pasta Origem " + c_pasta_origem)
print("Pasta Destino " + c_pasta_destino)
print(ordens)












#-Begin-----------------------------------------------------------------

#-Includes--------------------------------------------------------------
import sys, win32com.client

#-Sub Main--------------------------------------------------------------
def Anexar_SAP():

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
    session.findById("wnd[0]/tbar[0]/okcd").text = "iw38"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/chkDY_MAB").selected = -1
    session.findById("wnd[0]/usr/chkDY_MAB").setFocus()
    session.findById("wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[0]/usr/ctxtAUART-LOW").text = "*"
    session.findById("wnd[0]/usr/ctxtAUART-LOW").caretPosition = 1
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = "0105"
    session.findById("wnd[0]/usr/ctxtIWERK-LOW").caretPosition = 4
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtINGPR-LOW").text = "*"
    session.findById("wnd[0]/usr/ctxtINGPR-LOW").caretPosition = 1
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/ctxtVARIANT").text = "/ordem"
    session.findById("wnd[0]/usr/ctxtVARIANT").setFocus()
    session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 6
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(-1, "")
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll()
    session.findById("wnd[0]/tbar[1]/btn[42]").press()
    session.findById("wnd[0]/titl/shellcont/shell").pressButton("%GOS_TOOLBOX")
    session.findById("wnd[0]/shellcont/shell").pressContextButton("CREATE_ATTA")
    session.findById("wnd[0]/shellcont/shell").selectContextMenuItem("PCATTA_CREA")
    session.findById("wnd[1]").sendVKey(4)
    session.findById(
        "wnd[2]/usr/ctxtDY_PATH").text = "V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte Albuquerque\Desenvolvimento\Pedro\Pycharm\Anexar_Ordem_SAP\A Transferir"
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "2018283469_12-11-19_OPER-10_C.jpg"
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 33
    session.findById("wnd[2]").sendVKey(0)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
    session.findById("wnd[0]/tbar[0]/btn[15]").press()
    session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press()
  except:
    print(sys.exc_info()[0])

  finally:
    session = None
    connection = None
    application = None
    SapGuiAuto = None

#-Anexar_SAP------------------------------------------------------------------
Anexar_SAP()

#-End-------------------------------------------------------------------












