from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time
import openpyxl
import pandas as pd

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)
navegador.get("https://annas-archive.org/")

time.sleep(5)

navegador.find_element('xpath','/html/body/main/form/div/input').send_keys("Harry Potter")
navegador.find_element('xpath','/html/body/main/form/div/button').click()
navegador.find_element('xpath','/html/body/main/form/div[1]/select[1]').click()
navegador.find_element('xpath','/html/body/main/form/div[1]/select[1]/option[8]').click()
navegador.find_element('xpath','/html/body/main/form/div[2]/button').click()

time.sleep(5)

workbook = openpyxl.load_workbook('C:\\Users\\fellipi.borges\\PycharmProjects\\Robot\\Ticket.xlsx')#Entro na minha planilha
sheet = workbook.active #Informo que a planilha que quero é a Plan1(OpenpyXl entende que é a ultima tela aberta)
celula = sheet['A2'].value #Armazeno o valor da celula e digo que quero o valor dela
i = str(celula)  # Converto o dado de Int para Str para que possa ser concatenado posteriormente
print(i)  # Aqui eu imprimo o tipo de dado no qual eu estou trabalhando(só pra ter certeza que foi)

time.sleep(2)

j = navegador.find_element('xpath', '/html/body/div/div[2]/div[1]/a/h1').text #Puxo a informação do site e guardo ela em variavel
sheet['B2'] = j #Informo a celula onde quero armazenar minha informação e digo que ela está armazenada na variavel J
df = pd.read_excel('C:\\Users\\fellipi.borges\\PycharmProjects\\Robot\\Ticket.xlsx')#Leio meu arquivo Excel
df.at[2,'Resultado'] = j #Guardo a informação dentro da celula e digo que a informação tá na variavel J

workbook.save('C:\\Users\\fellipi.borges\\PycharmProjects\\Robot\\Ticket.xlsx') #Salvo a planilha com a informação armazenada

print(j)#quero ver o resultado da variavel só pra desencargo de consciencia

time.sleep(900)

