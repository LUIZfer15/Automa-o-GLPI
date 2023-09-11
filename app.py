import pandas as pd
import pyautogui
import time

tabela = pd.read_excel("dados.xlsx")
print(tabela)

# Abrir o navegador:
pyautogui.press('win')
time.sleep(1)
pyautogui.write('chrome')
time.sleep(1)
pyautogui.press('enter')
pyautogui.press('alt', 'tab')
time.sleep(1)
pyautogui.click(875,617)


# Acessar o GLPI:



#for i, nome in enumerate(tabela["Nome Colaborador"]):
    #marca = tabela.loc[i, "MarcaNote"]http://189.85.184.18/
    #modelo = tabela.loc[i, "ModeloNote"]
    #codigo = tabela.loc[i, "Codigo"]http://189.85.184.18/http://189.85.184.18/
    #patrimonio = tabela.loc[i, "Patrimonio"]
    #cargo = tabela.loc[i, "Cargo"]
    #processador = tabela.loc[i, "Processador"]
    #ram = tabela.loc[i, "Memoria"]
    #disco = tabela.loc[i, "DiscoRigido"]