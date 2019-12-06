
#Web scraper for purchases. Updates a database.

from selenium import webdriver
import time
from bs4 import BeautifulSoup
import pandas as pd

driver = webdriver.Chrome(r"C:\Users\bmvm\Downloads\Pedro\chrome driver\chromedriver.exe")

#ordem, pedido, reserva, rc, etapa_atual, NM, d_necc, PFF, Previsao = []

#ordem=[] #List to store ordens
#pedido=[] #List to store pedidos
#reserva=[] #List to store reservas
#rc=[] #List to store req. compra
#etapa_atual=[] #List to store pedidos
#nm=[]   #List to store NM de Materiais
#PFF=[] #List to store Previsão de Entrega pelo SAP
#previsao=[] #List to store Previsão de Entrega pelo localiza-e ou pedido


driver.get("https://localizae.petrobras.com.br/?q=4509675103")
time.sleep(5) # Let the user actually see something!
input("Feche todos os avisos PVF")


texto=[]
tcompleto=[]
content = driver.page_source
soup = BeautifulSoup(content)



for a in soup.findAll('timecard',href=True, attrs={'class':'display dataTable no-footer'}):
    texto=a.find('div', attrs={'class':'ul1'})
    tcompleto = a.find('div', attrs={'class': 'circ_completo'})
    #price=a.find('div', attrs={'class':'timecard'})
    #rating=a.find('div', attrs={'class':'hGSR34 _2beYZw'})
    texto.append(texto.text)
    tcompleto.append(tcompleto.text)
    #ratings.append(rating.text)