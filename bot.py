'''
- COMO AUTOMATIZAR ESSE PROCESSO?
- ONDE SERÁ FEITO? (VERSÃO WEB)

- QUAIS TECNOLOGIAS PREISO PARA RESOLVER ESSA DEMANDA:
1 - TECLADO (pyautogui)
2 - ACESSO AO SITE (webbrowser)
3 - AUTOMATIZAR DIGITAÇÃO (link whatsapp personalizado)
4 - AUTOMATIZAR LEITURA DE DADOS (planilha - openpyxl)
'''

'''DEMANDA:
PRECISO AUTOMATIZAR MINHAS MENSAGENS DE COBRANÇA EM DETERMINADO DIA COM CLIENTES COM VENCIMENTO DIFERENTE.'''
#1 DESCREVER OS PASSOS MANUAIS E DEPOIS TRANSFORMAR EM CÓDIGO
#2 LER PLANILHA E GUARDAR INFORMAÇÕES COMO 'NOME', 'TELEFONE' E 'VENCIMENTO'.
#3 CRIAR LINKS PERSONALIZADOS DO WHATSAPP E ENVIAR MENSAGENS PARA CADA CLIENTE.
#4 COM BASE NOS DADOS DA PLANILHA

# import selenium (usado para webscrape), substitui facilmente o pyautogui

# import openpyxl
# from urllib.parse import quote
# import webbrowser
# from time import sleep
# import pyautogui


# webbrowser.open('https://web.whatsapp.com/')
# sleep(10)


# #2 LER PLANILHA E GUARDAR INFORMAÇÕES COMO 'NOME', 'TELEFONE' E 'VENCIMENTO'.
# workbook = openpyxl.load_workbook('teste.xlsx')
# pagina_clientes = workbook['Folha1']

# for linha in pagina_clientes.iter_rows(min_row=2):
#     nome = linha[0].value #nome
#     telefone = linha[1].value # telefone
#     data = linha[2].value # data de vencimento
#     print(nome)
#     print(telefone)
#     print(data)
#     mensagem = f'Testando algumas alterações, ignore as proximas mensagens...'
#     # mensagem = f'Oii {nome}, hoje, no dia {data.strftime("%d/%m/%Y")} criei uma automação e esse é o teste do meu primeiro projeto :)!'
#     link_personalizado = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
#     webbrowser.open(link_personalizado)
#     sleep(6)
#     try:
#         # seta = []
#         # seta = seta.append(pyautogui.locateCenterOnScreen('seta.png'))
#         sleep(2)
#         pyautogui.click(1576, -73)
#         sleep(2)
#         pyautogui.hotkey('ctrl', 'w')
#         sleep(2)
#     except:
#         print(f'Não foi possível enviar mensagem para {nome}')
#         with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
#             arquivo.write(f'{nome} - {telefone}\n')
#     #3 CRIAR LINKS PERSONALIZADOS DO WHATSAPP E ENVIAR MENSAGENS PARA CADA CLIENTE.


    




