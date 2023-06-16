from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time
import openpyxl

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)
navegador.get(
    "*Url")

time.sleep(5)

navegador.find_element('xpath', '//*[@id="user_email"]').send_keys("*email")
navegador.find_element('xpath', '//*[@id="user_password"]').send_keys("*password")

time.sleep(5)

navegador.find_element('xpath', '//*[@id="sign-in-submit-button"]').click()

time.sleep(5)

#Uso operadores para chamar minha planilha Excel
workbook = openpyxl.load_workbook('C:\\Users\\fellipi.borges\\PycharmProjects\\Robot\\Ticket.xlsx')
sheet = workbook.active
celula = sheet['A2'].value
i = str(celula)  # Converto o dado de Int para Str para que possa ser concatenado posteriormente
print(i)  # Aqui eu imprimo o tipo de dado no qual eu estou trabalhando(só pra ter certeza que foi)

navegador.get("URL Concatenada" + i)  # Concateno o link do site com o 'Ticket'

time.sleep(5)

sheet = workbook.active
celula = sheet['B2'].value
j = navegador.find_element('xpath', '//*[@id="ember3058"]/div[1]/div/div/div').text
print(j)

time.sleep(900)
#workbook.save('Ticket.xlsx')

time.sleep(5)
