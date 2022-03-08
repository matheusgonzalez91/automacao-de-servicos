import time
import pyautogui
import pyperclip
import pandas as pd

#Entrar no sistema da empresa
'''hotkey é um tecla de atalho
por exemplo: ctrl + t'''

pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')

time.sleep(3)

pyautogui.write('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.press('enter')
time.sleep(3)

#Navegar no sistema, até encontrar a base de dados
pyautogui.click(x=343, y=299, clicks=2)

time.sleep(2)

#Exportar a base de vendas

pyautogui.click(x=345, y=380)
time.sleep(2)

pyautogui.click(x=1236, y=193)
time.sleep(2)

pyautogui.click(x=1010, y=591)
time.sleep(5)

#Calcular o faturamento da base de dados

tabela = pd.read_excel(r'C:\Users\Matheus Gonzalez\Downloads\Vendas - Dez.xlsx')

faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()

time.sleep(3)
#Enviar o email para a diretoria
#Entrar no email

pyautogui.hotkey('ctrl', 't')
time.sleep(3)

time.sleep(3)
pyautogui.write('https://mail.google.com/mail/u/0/#inbox')
pyautogui.press('enter')
time.sleep(4)

#Click no botão escrever
pyautogui.click(x=75, y=211)
time.sleep(2)

#Escrever destinatário
pyautogui.write('matheus.s.gonzalez10@gmail.com')
pyautogui.press('tab')
time.sleep(2)

#Escrever assunto
pyautogui.press('tab')
time.sleep(3)

pyperclip.copy('Relatório de vendas')
pyautogui.hotkey('ctrl', 'v')
time.sleep(3)

#Escrever corpo do email
pyautogui.press('tab')
time.sleep(3)

texto = f'''
Olá, bom dia!

O faturamento de ontem foi de: R$ {faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Abs
Matheus Gonzalez

'''
pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')
time.sleep(3)

#Enviar email
pyautogui.hotkey('ctrl', 'enter')
