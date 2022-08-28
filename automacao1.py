import pyautogui
import pyperclip
import time
import pandas as pd


#Fazer o pyautogui dar uma descansada para executar as tarefas corretamente
pyautogui.PAUSE = 1

#############################PASSO 1: entrar no sistema(No caso no link do drive)###############################################
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
time.sleep(4)#demora um tempinho para carregar
pyautogui.click(x=774, y=405, clicks=2)
#####se o google tiver aberto já ele não ira abrir ;( #####
####Abaixo precisa ser o pyperclipe se não da erro devido aos caracteres especiais
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga") 
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")

time.sleep(3)

##############################PASSO 2: navegar no sistema e encontrar a base de dados(no caso a pasta "exportar")#############################
pyautogui.click(x=404, y=269, clicks=2)
time.sleep(4)


############################## PASSO 3: fazer o dowload da base de dados #############################
pyautogui.click(x=404, y=269)#arquivo
pyautogui.click(x=1165, y=155)#3 pontinhos
pyautogui.click(x=976, y=559)#baixar

time.sleep(7)#esperar baixar o arquivo


############################## PASSO 4: importar a base dados para o python #############################
tabela = pd.read_excel(r"C:\Users\Lenovo\Downloads\Vendas - Dez.xlsx")
print(tabela)


############################## Passo 5: fazer os calculos necessários #############################
#faturamento
faturamento = tabela["Valor Final"].sum()
print(faturamento)
#quantidade de produtos
quantidade = tabela["Quantidade"].sum()
print(quantidade)


############################## Passo 6: enviar os calculos #############################
#abrir uma nova guia
pyautogui.hotkey("ctrl","t")

#entrar no email
pyperclip.copy("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(2)

#clicar no escrever
pyautogui.click(x=121, y=164)
time.sleep(2)

#escrever o email destinatário
pyautogui.write("estevanmartins03@gmail.com")
pyautogui.press("tab") #seleciona destinatario
pyautogui.press("tab")#passsa paara o campo dos assuntos

#escrever o assunto
pyautogui.write("Relatório de vendas")
pyautogui.press("tab")#passa para o campo do corpo do email

#escrever o corpo do email
texto = f"""Olá, bom dia

Segue o faturamento: R${faturamento:}
Segue a quantidade de produtos vendidos: {quantidade:}

abraços,
Estevan Sousa"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")

#enviar
pyautogui.hotkey("ctrl","enter")


