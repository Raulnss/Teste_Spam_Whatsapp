import sys
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

# Determina o caminho do arquivo de forma apropriada
if getattr(sys, 'frozen', False):
    # Se o aplicativo está sendo executado como um executável
    base_path = os.path.dirname(sys.executable)
else:
    # Se o aplicativo está sendo executado como um script
    base_path = os.path.dirname(os.path.abspath(__file__))

# Construindo o caminho do arquivo
file_path = os.path.join(base_path, 'numeros.xlsx')

# Verificando se o arquivo existe
if not os.path.exists(file_path):
    print(f"Arquivo não encontrado: {file_path}")
    sys.exit()

workbook = openpyxl.load_workbook(file_path)
pagina_clientes = workbook['Plan1']

for linha in pagina_clientes.iter_rows(min_row=1):
    telefone = linha[0].value
    sequencia_str = str(telefone)
    if sequencia_str[2] == '1' and sequencia_str[3] == '3':
        mensagem = 'Spam de whatsapp'
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        pyautogui.press('enter')
        sleep(2)
        pyautogui.hotkey('ctrl', 'w')
        sleep(2)
    else:
        print("As condições não foram atendidas.")
