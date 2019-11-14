
# -Begin-----------------------------------------------------------------

# -Includes--------------------------------------------------------------
import sys
import win32com.client
import os
import datetime
import xlwings as xw
import easygui
import pandas as pd


# Selecionar pasta de origem









c_pasta = r"V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte "\
          r"Albuquerque\Desenvolvimento\Pedro\Pycharm\Anexar_Ordem_SAP\A Transferir"
n_file = "Ordem" + datetime.datetime.today().strftime("%Y%m%d") + ".xlsx"
c_file = os.path.join(c_pasta, n_file)

# Abrir fonte dos dados das ordens (https://stackoverflow.com/questions/17977540/pandas-looking-up-the-list-of-sheets-in-an-excel-file)
path_fonte_dados = easygui.fileopenbox()
xl = pd.ExcelFile(path_fonte_dados)
aba = easygui.buttonbox('Escolha a aba a carregar', 'Abas:', xl.sheet_names)
dados = xl.parse(aba, skiprows=18)
dados['OM'].to_clipboard(excel=True, sep=None, index=False, header=False)






#-Sub Main--------------------------------------------------------------
def Main():

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

    if os.path.exists(c_file):
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        print('Substituido arquivo em ' + c_file)
    else:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        print("criado arquivo em " + c_file)



  except:
    print(sys.exc_info()[0])

  finally:
    session = None
    connection = None
    application = None
    SapGuiAuto = None

#-Main------------------------------------------------------------------
Main()

#-End-------------------------------------------------------------------


a = os.listdir(r"V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte "
               r"Albuquerque\Desenvolvimento\Pedro\Pycharm\Anexar_Ordem_SAP\A Transferir")

c = list(range(len(a)))

for f in range(0, len(a)):
    low = str.find(a[f], " - ") + 3
    upper = str.find(a[f], "_")
    b = a[f]
    c[f] = b[low:upper]
    print(c, low, upper)

Ordens = c
