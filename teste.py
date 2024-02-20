'''import pyautogui
from time import sleep

sleep(5)
x, y = pyautogui.position()
print(x, y)
IDENTIFICAR E RETORNAR NO TERMINAL A POSIÇÃO X, Y DO MOUSE!
'''

from funcionalidades import *

def main():
    # Executa o processo de leitura da planilha, envio de mensagens e registro de erros
    contatos = ler_arquivo('planilha.xlsx')
    erros = []
    abrir_wpp()
    for contato in contatos:
        try:
            enviar_msg(contato)
        except Exception as e:  # Deve-se capturar exceções específicas
            print(f"Não foi possível enviar mensagem para {contato[0]}")
            erros.append((contato[0], contato[1]))
    if erros:
        registro_erros(erros)

main()