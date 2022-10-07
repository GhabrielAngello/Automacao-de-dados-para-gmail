import pyautogui
import pyperclip
import time
import pandas

arquivo = r'C:\Users\PC\Downloads\Vendas - Dez.xlsx'
tabela = pandas.read_excel(r"C:\Users\PC\Downloads\Vendas - Dez.xlsx")

pyautogui.PAUSE = 1
#passo 1
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(2)

#passo 2
while not pyautogui.locateCenterOnScreen("exp.PNG", confidence = 0.9):
    time.sleep(2)
#passo 3
pyautogui.click(x=324, y=295, clicks=2)

while not pyautogui.locateOnScreen("vendas.PNG", confidence = 0.9):
    time.sleep(2)

pyautogui.click(x=353, y=355)
pyautogui.click(x=1154, y=189)
pyautogui.click(x=937, y=561)
time.sleep(3)

#passo 4
print(tabela)
#passo 5
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

#passo 5
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/?tab=rm&ogbl")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

while not pyautogui.locateOnScreen("escrever.PNG"):
    time.sleep(1)

pyautogui.click("escrever.PNG")
time.sleep(2)
pyautogui.write("pcghabriel@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

texto = f"""
Prezados,

Segue Relatório de Vendas
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}

Estou à disposição
Att, Ghabriel Angello
"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# passo 6

pyautogui.click("anexo.PNG")
pyperclip.copy(arquivo)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
while not pyautogui.locateOnScreen("finalizado.PNG"):
    time.sleep(2)
#ultimo passo 7
pyautogui.hotkey("ctrl", "enter")