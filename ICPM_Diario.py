
# -Begin-----------------------------------------------------------------

# -Includes--------------------------------------------------------------
import sys
import win32com.client
import os
import datetime
import xlwings as xw
import easygui
import pandas as pd


# -Caminho da pasta de Salvamento V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte
# Albuquerque\Desenvolvimento\Pedro\Pycharm\ICPM Diário\Parcial

c_pasta = r"V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte " \
          r"Albuquerque\Desenvolvimento\Pedro\Pycharm\ICPM Diário\Parcial"
n_file = "ICPM_" + datetime.datetime.today().strftime("%Y%m%d") + ".xlsx"
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
    session.findById("wnd[0]/tbar[0]/okcd").text = "iw37n"
    session.findById("wnd[0]").sendVKey(0)
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    session.findById("wnd[0]").sendVKey(0)
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").select()
    session.findById("wnd[0]/usr/chkSP_MAB").selected = -1
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = "/MD_progsem"
    session.findById("wnd[0]/usr/chkSP_MAB").setFocus()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow = 2
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
    session.findById("wnd[1]").sendVKey(4)
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = c_pasta
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = n_file
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 18
    session.findById("wnd[2]/tbar[0]/btn[11]").press()
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


def compara_excel():
  import os
  import win32com.client as win32

  excel = win32.Dispatch("Excel.Application")
  # DispatchEx creates a new instance,
  # while Dispatch uses an existing one if one exists.

  book_name = n_file
  book_path = c_pasta + "\\" + book_name

  wb = excel.Workbooks.Open(book_path)
  ws = wb.Worksheets(1)

  ws.Columns("A:A").Select
  excel.Selection.Insert
  excel.Shift = xlToRight
  excel.Range("A2").Select
  ws.ActiveCell.FormulaR1C1 = "=RC[3]&""/""&VALUE(RC[4])"
  excel.Range("A2").Select
  excel.Selection.AutoFill
  ws.Destination := Range("A2:A72")
  Range("A2:A72").Select

  msoShapeOval = 9
  ws.Shapes.AddShape(msoShapeOval, 270.75, 205.5, 10.0, 10.0).Select()
  excel.Selection.ShapeRange.Fill.ForeColor.RGB = 255

  wb.SaveAs(os.getcwd() + '/' + 'line-chart-2.xlsx')
  excel.Quit()




  compara_excel()


