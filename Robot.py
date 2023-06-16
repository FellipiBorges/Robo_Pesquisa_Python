from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time
import openpyxl
import requests
from lxml import html

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)
navegador.get("#Aqui vai a URL")
navegador.find_element('xpath','//*[@id="user_email"]').send_keys("#E-mail")
navegador.find_element('xpath','//*[@id="user_password"]').send_keys("#Password")
navegador.find_element('xpath','//*[@id="sign-in-submit-button"]').click()


workbook = openpyxl.load_workbook('C:\\Users\\fellipi.borges\\PycharmProjects\\Robot\\Ticket.xlsx') #Uma planilha dentro da pasta do projeto que contem um ID a ser pesquisado na URL
sheet = workbook.active
celula = sheet['A2'].value
i = str(celula) #Converto o dado de Int para Str para que possa ser concatenado posteriormente
print((i)) #Aqui eu imprimo o tipo de dado no qual eu estou trabalhando(só pra ter certeza que foi)

navegador.get("URL concatenada com o ID da planilha" + i) #Concateno o link do site com o 'Ticket'


book.save('Ticket.xlsx')





