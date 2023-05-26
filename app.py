"""Automação GLPI."""

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pandas as pd
import pyautogui
import keyboard
import time

servico = Service(ChromeDriverManager().install())

navegador = webdriver.Chrome(service=servico)

tabela = pd.read_excel("PATRIMÔNIO EM USO .xlsx")
print(tabela)

# Abrir navegador:
navegador.get("http://189.85.184.18/")

# Logar no GLPI:
navegador.find_element('xpath', '//*[@id="login_name"]').send_keys("luiz.mendes")
time.sleep(1)
navegador.find_element('xpath', '/html/body/div[1]/div/div/div[2]/div/form/div/div/div[3]/input').send_keys("So5Bi5Ri1")
time.sleep(1)
navegador.find_element('xpath', '/html/body/div[1]/div/div/div[2]/div/form/div/div/div[6]/button').click()
time.sleep(2)

# Colocar em tela cheia:
pyautogui.click(878,21)
time.sleep(2)

# Clicar em ativo:
navegador.find_element('xpath', "//span[normalize-space()='Ativos']").click()
time.sleep(1)

# Clicar em computadores:
navegador.find_element('xpath', "//a[@title='Computadores']//span[@class='text-wrap']").click()
time.sleep(1)

# Clicar em adicionar:
navegador.find_element('xpath', "//span[normalize-space()='Adicionar']").click()
time.sleep(2)

for i, nome in enumerate(tabela["Nome Colaborador"]):
    marca = tabela.loc[i, "Marcanote"]
    modelo =tabela.loc[i, "ModeloNote"]
    codigo = tabela.loc[i, "Codigo"]
    patrimonio = tabela.loc[i, "PatrimonioNote"]
    cargo = tabela.loc[i, "Cargo"]
    processador = tabela.loc[i, "Processador"]
    ram = tabela.loc[i, "Memoria"]
    disco = tabela.loc[i, "DiscoRigido"]

    # Adicionar nome:
    navegador.find_element('xpath', "//input[@id='name_1113727539']").send_keys(marca)
    keyboard.press('space')
    keyboard.write(modelo)

    # Adicionar localização:
    navegador.find_element('xpath', "//span[@id='select2-dropdown_locations_id892556012-container']").click()
    navegador.find_element('xpath', "//div[@id='add_dropdown_autoupdatesystems_id211934654']").click()

    # Adicionar usuário:
    navegador.find_element('xpath', "//div[@id='add_dropdown_autoupdatesystems_id211934654']").send_keys(nome)
    keyboard.press("enter")

    # Adicionar grupo:
    navegador.find_element('xpath', "//span[@id='select2-dropdown_groups_id1364875096-container']").send_keys(cargo)
    keyboard.press("enter")

    # Adicionar status:
    pyautogui.click(1547,352)
    time.sleep(1)
    pyautogui.click(1786,563)
    time.sleep(1)
    pyautogui.click(1531,516)
    time.sleep(1)

    # Adicionar o tipo de computador:
    navegador.find_element('xpath', "//span[@id='select2-dropdown_computertypes_id1005427307-container']").click()
    navegador.find_element('xpath', "//li[@title='Notebook - ']").click()

    # Adicionar fabricante:
    navegador.find_element('xpath', "//span[@id='select2-dropdown_manufacturers_id1048659121-container']").send_keys(marca)
    keyboard.press("enter")

    # Adicionar modelo:
    navegador.find_element('xpath', "//span[@id='select2-dropdown_computermodels_id1989611248-container']").send_keys(modelo)
    keyboard.press("enter")

    # Adicionar n° de série:
    navegador.find_element('xpath', "//input[@id='serial_928891608']").send_keys(codigo)

    # Adicionar patrímonio:
    navegador.find_element('xpath', "//input[@id='otherserial_1538519314']").send_keys(patrimonio)

    # Adicionar notebook:
    navegador.find_element('xpath', "//button[@name='add']//span[contains(text(),'Adicionar')]").click()
    time.sleep(2)

    # Clicar em computadores;
    navegador.find_element('xpath', "//a[@class='here']").click()

    # Abrir o ativo salvo:
    navegador.find_element('xpath',"//a[@class='here']").send_keys(marca)
    keyboard.press("space")
    keyboard.write(modelo)
    keyboard.press('enter')
    navegador.find_element('xpath', "//a[@id='Computer_133_133']").click()
    time.sleep(2)

    # Clicar em sistemas operacionais:
    navegador.find_element('xpath', "//a[normalize-space()='Sistemas operacionais']").click()
    navegador.find_element('xpath', "//span[@id='select2-dropdown_operatingsystems_id4322937-container']").click()
    navegador.find_element('xpath', "//li[@title='Windows - ']").click()

    # Adicionar sistema:
    navegador.find_element('xpath', "//button[@name='add']//span[contains(text(),'Adicionar')]").click()

    # Adicionar componentes;
    navegador.find_element('xpath', "//a[contains(text(),'Componentes')]").click()

    # Adionar processador?
    navegador.find_element('xpath', "//span[@id='select2-dropdown_devicetype1158850106-container']//span[contains(text(),'-----')]").send_keys("processador")
    keyboard.press("enter")
    navegador.find_element('xpath', "//span[@id='select2-dropdown_devices_id1158850106-container']").send_keys(processador)
    keyboard.press("enter")
    navegador.find_element('xpath', "//span[@aria-expanded='true']//span[@role='presentation']").send_keys("1")
    keyboard.press("enter")
    navegador.find_element('xpath', "//input[@name='add']").click()
    time.sleep(3)

    # Adicionar Memoria Ram:
    navegador.find_element('xpath', "//span[@id='select2-dropdown_devicetype1158850106-container']//span[contains(text(),'-----')]").send_keys("memoria")
    keyboard.press("enter")
    navegador.find_element('xpath', "//span[@id='select2-dropdown_devices_id1158850106-container']").send_keys(ram)
    keyboard.press("enter")
    navegador.find_element('xpath', "//span[@aria-expanded='true']//span[@role='presentation']").send_keys("1")
    keyboard.press("enter")
    navegador.find_element('xpath', "//input[@name='add']").click()
    time.sleep(3)

    # Adicionar Disco Rigido:
    navegador.find_element('xpath', "//span[@id='select2-dropdown_devicetype1158850106-container']//span[contains(text(),'-----')]").send_keys("disco")
    keyboard.press("enter")
    navegador.find_element('xpath', "//span[@id='select2-dropdown_devices_id1158850106-container']").send_keys(disco)
    keyboard.press("enter")
    navegador.find_element('xpath', "//span[@aria-expanded='true']//span[@role='presentation']").send_keys("1")
    keyboard.press("enter")
    navegador.find_element('xpath', "//input[@name='add']").click()
    time.sleep(3)