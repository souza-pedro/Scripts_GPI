import sys
import easygui as g
import os
import pandas as pd


# Escolha da pasta de Destino

def escolher_pasta():
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

        if choice == "Escolher":
            choice = g.diropenbox()
            c_pasta_origem = choice

        # g.msgbox("Você escolheu: " + str(c_pasta_origem), "Resultado Escolha")

        msg = "Gostaria de Continuar?   Você escolheu: " + str(c_pasta_origem)
        title = "Please Confirm"
        if g.ccbox(msg, title):  # show a Continue/Cancel dialog
            pass  # user chose Continue
            break
        else:
            choice = ""  # user chose cancel
    # user chose Cancel

    # Escolha da pasta de Saída

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

        if choice == "Escolher":
            choice = g.diropenbox()
            c_pasta_destino = choice

        # g.msgbox("Você escolheu: " + str(c_pasta_destino), "Resultado Escolha")

        msg = "Gostaria de Continuar?   Você escolheu: " + str(c_pasta_destino)
        title = "Please Confirm"
        if g.ccbox(msg, title):  # show a Continue/Cancel dialog
            pass  # user chose Continue
            return c_pasta_origem, c_pasta_destino
            # break
        else:
            choice = ""  # user chose cancel


def lista_ordens_clipboard(lista, retorno):
    # Transformando Nomes dos aquivos em lista de Ordens.
    # Nome do arquivo deve estar no formato XXXXXXXXXX_xx-xx-xx_OPER-XX_C (Nº Ordem_dia-mes_ano_OPER-Nº Oper_C)

    a = os.listdir(lista)

    c = list(range(len(a)))

    for f in range(0, len(a)):
        # low = str.find(a[f], " - ") + 3
        low = 0
        upper = str.find(a[f], "_")
        b = a[f]
        c[f] = b[low:upper]
        # print(c, low, upper)

    # Ordens = c
    ordens = pd.DataFrame(c)
    #ordens.to_clipboard(index=False, header=False)

    return lista, ordens


def copiar_destino(c_origem, c_destino, arquivo):
    import shutil
    c_arquivo = os.path.join(c_origem, arquivo)
    shutil.copy(c_arquivo, c_destino)
    c_arquivo_dest = os.path.join(c_destino, arquivo)
    novonome = arquivo[:str.find(arquivo, ".") - 1] + "_OK" + arquivo[str.find(arquivo, "."):]
    os.rename(c_arquivo_dest, novonome)
    print("Copiado para OK arquivo " + novonome)


# -Sub anexar SAP--------------------------------------------------------------
def anexar_sap(c_origem, c_destino, ordens):
    import sys
    import win32com.client

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

        # >Insert your SAP GUI Scripting code here<
        # Pela IW33
        # session.findById("wnd[0]").resizeWorkingPane(95, 25, 0)
        session.findById("wnd[0]/tbar[0]/okcd").text = "iw33"
        session.findById("wnd[0]").sendVKey(0)

        for i in range(0, len(ordens)):
            ordem = ordens[0]
            session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = ordem[i]
            print("Abriu ordem " + ordem[i])
            session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").caretPosition = 10
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/titl/shellcont/shell").pressButton("%GOS_TOOLBOX")
            session.findById("wnd[0]/shellcont/shell").pressContextButton("CREATE_ATTA")
            session.findById("wnd[0]/shellcont/shell").selectContextMenuItem("PCATTA_CREA")
            lista = os.listdir(c_origem)

            session.findById("wnd[1]/usr/ctxtDY_PATH").text = c_origem
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = lista[i]
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 33
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[0]/shellcont").close()
            print("Anexou arquivo " + lista[i])
            # session.findById("wnd[0]/tbar[0]/btn[3]").press() #voltar
            try:
                session.findById("wnd[0]/tbar[0]/btn[11]").press()  # Salvar
                # print("tentou sair")
                session.findById("wnd[0]/tbar[0]/btn[3]").press()  # voltar
                # print("tentou voltar")
            except:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
                print("voltou simples")
            copiar_destino(c_origem, c_destino, f_nome)

        session.findById("wnd[0]/tbar[0]/btn[3]").press()  # voltar

    except:
        print(sys.exc_info()[0])

    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None


def main():
    # Escolhe a pasta padrão
    (c_origem, c_destino) = escolher_pasta()
    print("Pasta de Origem " + c_origem)
    print("Pasta destino " + c_destino)
    print("____________________________________________")

    # Copia Nº de Ordens da pasta para o clipboard
    (c_origem, ordens) = lista_ordens_clipboard(c_origem, "")
    # print("C_origem " + c_origem)
    print("Lista de Ordens:")
    print(ordens)
    print("____________________________________________")

    # Anexar no SAP
    anexar_sap(c_origem, c_destino, ordens)

    # print(c_origem)
    print("Fim do programa")
    # print("Pasta de Origem " + c_origem)
    # print("Pasta destino " + c_destino)


main()
