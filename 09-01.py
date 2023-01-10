import pyautogui as pag
import time
import pyperclip as ppc
import pandas as pd
pag.PAUSE = 1

pag.press('win')
pag.write('opera')
pag.press('enter')
ppc.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pag.hotkey('ctrl', 'v')
pag.press('enter')

time.sleep(3)

pag.click(x=425, y=306, clicks=2)

time.sleep(2)

pag.click(x=425, y=306)
pag.click(x=1161, y=190)
pag.click(x=999, y=589)

tabela = pd.read_excel(r"C:\Users\Matheus\Downloads\Vendas - Dez.xlsx")

faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

pag.hotkey('ctrl', 't')
pag.click(x=945, y=678)
time.sleep(3)
pag.click(x=123, y=200)
time.sleep(3)
ppc.copy(r"matheusfelippe200@gmail.com")
pag.hotkey('ctrl', 'v')
pag.press('tab')
pag.write("F & Q")
pag.press('tab')

texto = f"""O faturamento total foi de R${faturamento:,.2f}
e a quantidade de vendas foi de: {quantidade:,}. """

ppc.copy(texto)
pag.hotkey('ctrl', 'v')
pag.hotkey('ctrl', 'enter')