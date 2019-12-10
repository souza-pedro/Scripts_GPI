
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


products=[] #List to store name of the product
prices=[] #List to store price of the product
ratings=[] #List to store rating of the product

content = driver.page_source
soup = BeautifulSoup(content, 'html.parser')

print(soup.title)
for a in soup.find_all("div", attrs={'class':'jsonHide'}):
    products.append(a.text)

b = products[0]
requisicao = b[b.find('requisicao')+13:b.find('requisicao')+21]
entrega = b[b.find('"entrega"')+11:b.find('"entrega"')+11+8]

