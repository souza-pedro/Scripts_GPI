import os
import pandas as pd
import win32com.client
import fnmatch
import glob


p_salvamento = r"V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte Albuquerque\Desenv\Pedro\Pycharm\Confirmação Semanal"
c_programação = r"V:\COMPARTILHADO_CSC-SSE_NSIF\NP-1\GPI\1 - Programações Semanais\INDUSTRIAL\2019"
L = []
c_sem = []

#with os.scandir(c_programação) as entries:
#    for entry in entries:
#        print(entry.name)
#        L.append(entry.name)

sem = input("digite a semana para análise")

sem_string = c_programação + r"\*" + sem + r"*"

c_sem = glob.glob(sem_string)



for filename in os.listdir(c_sem[0]):
    if fnmatch.fnmatch(filename, "*oficial*.xlsm"):
        L.append(filename)
        print(filename)
print(len(c_sem))


