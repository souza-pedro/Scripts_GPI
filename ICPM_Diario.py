
#-Begin-----------------------------------------------------------------

#-Includes--------------------------------------------------------------
import sys
import win32com.client
import os
import datetime
import xlwings as xw


# -Caminho da pasta de Salvamento
#V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte Albuquerque\Desenvolvimento\Pedro\Pycharm\ICPM Diário\Parcial
c_pasta = r"V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte " \
          r"Albuquerque\Desenvolvimento\Pedro\Pycharm\ICPM Diário\Parcial"
n_file = "ICPM_" + datetime.datetime.today().strftime("%Y%m%d") + ".xlsx"
c_file = os.path.join(c_pasta, n_file)

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

    # Editar o Excel

    # Start by opening the spreadsheet and selecting the main sheet
    workbook = load_workbook(filename=c_file)
    sheet = workbook.active

    # Write what you want into a specific cell
    sheet.insert_cols(idx=1)


    sheet["C1"] = "writing ;)"

    # Save the spreadsheet
    workbook.save(filename=c_file)

    xw.books.open(c_file)
    xw.Range("A:A").select()
    xw.
    xw.Shift := xlToRight
    xw.Range("A2").Select
    xw.ActiveCell.FormulaR1C1 = "=RC[3]&""/""&VALUE(RC[4])"
    xw.Range("A2").Select
    xw.Selection.AutoFill
    xw.Destination := Range("A2:A70")
    xw.Range("A2:A70").Select

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







#Macro Excel

    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=RC[3]&""/""&VALUE(RC[4])"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A70")
    Range("A2:A70").Select



# Texto do SAp


# session.findById("wnd[0]/tbar[0]/okcd").text = "iw37n"
# session.findById("wnd[0]").sendVKey(0)
# session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
# session.findById("wnd[0]").sendVKey(0)
# session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press()
# session.findById("wnd[1]/tbar[0]/btn[24]").press()
# session.findById("wnd[1]/tbar[0]/btn[8]").press()
# session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").select()
# session.findById("wnd[0]/usr/chkSP_MAB").selected = -1
# session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = "/MD_progsem"
# session.findById("wnd[0]/usr/chkSP_MAB").setFocus()
# session.findById("wnd[0]/tbar[1]/btn[8]").press()
# session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow = 2
# session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll()
# session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
# session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
# session.findById("wnd[1]/tbar[0]/btn[0]").press()
# session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
# session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
# session.findById("wnd[1]").sendVKey(4)
# session.findById("wnd[2]/usr/ctxtDY_PATH").text = "V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte Albuquerque\Desenvolvimento\Pedro\Pycharm\ICPM Diário\Parcial"
# session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "ICPM_20191106.xlsx"
# session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 18
# session.findById("wnd[2]/tbar[0]/btn[11]").press()
# session.findById("wnd[1]/tbar[0]/btn[0]").press()
