import os
import pandas as pd
import win32com.client
import fnmatch
import glob
import easygui as g

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
dados = pd.read_excel(c_arquivo, skiprows=5, sheet_name='Programação Semanal', convert_float=True)


