# -Begin-----------------------------------------------------------------

# -Includes--------------------------------------------------------------
import sys
import win32com.client
from win32com.client import Dispatch
import datetime
import os
import pandas as pd
import numpy as np

# -Caminho da pasta de Salvamento
c_pasta = r"V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte Albuquerque\Desenvolvimento" \
          r"\Pedro\Pycharm\IARI Diário"
nome_file = datetime.datetime.today().strftime("%y%m%d") + ".xlsx"
c_arquivo = os.path.join(c_pasta, nome_file)


# -Sub Main--------------------------------------------------------------
def main():


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
    # session.findById("wnd[0]").maximize()
    nome_file = datetime.datetime.today().strftime("%y%m%d") + ".xlsx"
    c_arquivo = os.path.join(c_pasta, nome_file)
    session.findById("wnd[0]/tbar[0]/okcd").text = "iw67"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtQMART-LOW").text = "zr"
    session.findById("wnd[0]/usr/ctxtQMART-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtQMART-LOW").caretPosition = 2
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = "0105"
    session.findById("wnd[0]/usr/ctxtIWERK-LOW").caretPosition = 4
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    session.findById("wnd[0]/usr/ctxtDATUB").setFocus()
    session.findById("wnd[0]/usr/ctxtDATUB").caretPosition = 0
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "PARNR"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = c_pasta
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_file

    if os.path.exists(c_arquivo):
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        print('Substituido arquivo em ' + c_arquivo)
    else:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        print("criado arquivo em " + c_arquivo)


    input("Continuar irá fechar todas as instâncias do excel...")

    # Find Excel session and closes ALL OF THEM (https://stackoverflow.com/questions/6337595/python-win32-com-closing
    # -excel-workbook)
    excel = Dispatch("Excel.Application")
    excel.Visible = False

    # without saving
    map(lambda book: book.Close(False), excel.Workbooks)
    excel.Quit()

    print(c_arquivo)

    # format_excel(c_arquivo, c_pasta)
    # Início do tratamento do excel
    dados = pd.read_excel(c_arquivo)
    dados.info()
    dados['tipo_nota'] = np.where(dados['Código de medidas'] == "D", pd.DateOffset(days=540),
                                  np.where(dados['Código de medidas'] == "C", pd.DateOffset(days=360),
                                           np.where(dados['Código de medidas'] == "B", pd.DateOffset(days=120),
                                                    np.where(dados['Código de medidas'] == "A",
                                                             pd.DateOffset(days=15),
                                                             False))))
    dados['Prazo_medida'] = dados['Data de criação'] + dados['tipo_nota']
    dados['No_prazo?'] = np.where(dados['Prazo_medida'] >= pd.Timestamp('today'), "OK", "Em Atraso")
    nome_file = "IARI Diário " + datetime.datetime.today().strftime("%y-%m-%d") + ".xlsx"
    c_file = os.path.join(c_pasta, nome_file)
    # Correções -  formatar data de criação, centro de localização como texto,  formatar Prazo_medida
    # Realizados: tirar coluna tipo_nota, tirar index,
    dados = dados.drop('tipo_nota', axis=1)
    dados.to_excel(c_file, index= False)





# -Main------------------------------------------------------------------
main()


# -End--------------------------