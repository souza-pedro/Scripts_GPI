import sys
import easygui as g


# Escolha da pasta de Destino

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
         r"      Pasta Padrão: " + c_pasta_origem
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

print("FIM " + c_pasta_origem + c_pasta_destino)