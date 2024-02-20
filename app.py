import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

webbrowser.open('https://web.whatsapp.com/')
sleep(15)

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Planilha1']


for linha in pagina_clientes.iter_rows(min_row=2):
    if linha.value == '':
        break
    else:
        nome = linha[0].value
        telefone = linha[1].value

        mensagem = f'Aooba {nome}, isso é só um teste de mensagem automatizada'

        try:
            link_mensagem = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
            webbrowser.open(link_mensagem)
            sleep(10)
            seta = pyautogui.locateCenterOnScreen('seta.png')
            sleep(4)
            pyautogui.click(seta[0], seta[1])
            sleep(4)
            pyautogui.hotkey('ctrl', 'w')
            sleep(4)
        except:
            print(f'Não foi possível enviar mensagem para {nome}')
            with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                arquivo.write(f'{nome}, {telefone}{os.linesep}')
