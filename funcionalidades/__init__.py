import openpyxl
from urllib.parse import quote
import webbrowser
import pyautogui
from time import sleep # Substituir por outra função

def abrir_wpp(): 
    # Abre o WhatsApp Web no navegador padrão
    webbrowser.open('https://web.whatsapp.com/') 
    sleep(5) 
    # Espera o WhatsApp carregar


def ler_arquivo(arquivo): 
    # Lê os dados de uma planilha Excel
    workbook = openpyxl.load_workbook(arquivo)
    # pagina = workbook['Pagina1']
    # return [(linha[0].value, linha[1].value, linha[2].value) for linha in pagina.iter_rows(min_row=2)]
    print('Deu certo')
    # Retorna uma lista de tuplas com os dados 'nome', 'telefone', 'data de vencimento'

def enviar_msg(contato):
    # Envia uma mensagem personalizada para um contato do WhatsApp
    nome, telefone, data = contato
    msg_personalizada = f"Oi {nome}, hoje, no dia {data.strftime('%d/%m/%Y')}, criei uma automação e esse é o teste do meu primeiro projeto! :)"
    link = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(msg_personalizada)}'
    webbrowser.open(link)
    sleep(5)  # Espera a mensagem estar pronta para enviar
    # Aqui você deveria ter um método mais confiável para enviar a mensagem
    # como a API do WhatsApp, ou checando por confirmação visual ou de status


def registro_erros(erros):
    # Escreve os erros em um arquivo CSV
    with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
        for nome, telefone in erros:
            arquivo.write(f'{nome} - {telefone}\n')

def main():
    # Executa o processo de leitura da planilha, envio de mensagens e registro de erros
    contatos = ler_arquivo('teste.xlsx')
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