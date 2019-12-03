
#-Includes--------------------------------------------------------------
import sys, win32com.client

#-Sub anexar SAP--------------------------------------------------------------
def anexar_sap(c_origem, ordens):

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
    # Pela IW33
    # session.findById("wnd[0]").resizeWorkingPane(95, 25, 0)
    session.findById("wnd[0]/tbar[0]/okcd").text = "iw33"
    session.findById("wnd[0]").sendVKey(0)
    for f_name in os.listdir(c_origem):
        if f_name.startswith('2018283469'):
            print(f_name)
    session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = "2018283469"
    session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").caretPosition = 10
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/titl/shellcont/shell").pressButton("%GOS_TOOLBOX")
    session.findById("wnd[0]/shellcont/shell").pressContextButton("CREATE_ATTA")
    session.findById("wnd[0]/shellcont/shell").selectContextMenuItem("PCATTA_CREA")
    session.findById(
        "wnd[1]/usr/ctxtDY_PATH").text = r"V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte Albuquerque\Desenvolvimento\Pedro\Pycharm\Anexar_Ordem_SAP\A Transferir"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "2018283469_12-11-19_OPER-10_C.jpg"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 33
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/shellcont").close()
    # session.findById("wnd[0]/tbar[0]/btn[3]").press() #voltar
    try:
        session.findById("wnd[0]/tbar[0]/btn[11]").press()  # Salvar
        print("tentou sair")
        session.findById("wnd[0]/tbar[0]/btn[3]").press()   # voltar
        print("tentou voltar")
    except:
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        print("voltou simples")

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


