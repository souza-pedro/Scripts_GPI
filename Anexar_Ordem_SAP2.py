import sys
import easygui as g

while 1:

    msg ="Gostaria de Selecionar a pasta ou usar a pasta padrão?" \
         r"      Pasta Padrão: V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte " \
         r"Albuquerque\Desenvolvimento\Pedro\Pycharm\Anexar_Ordem_SAP\A Transferir "
    title = "Escolha Pasta Origem"
    choices = ["Escolher", "Usar Padrão"]
    choice = g.choicebox(msg, title, choices)

    # note that we convert choice to string, in case
    # the user cancelled the choice, and we got None.
    g.msgbox("Você escolheu: " + str(choice), "Survey Result")

    msg = "Do you want to continue?"
    title = "Please Confirm"
    if g.ccbox(msg, title):     # show a Continue/Cancel dialog
        pass  # user chose Continue
    else:
        sys.exit(0)           # user chose Cancel