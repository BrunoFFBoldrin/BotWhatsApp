#Automatização WhatsApp - envio de cobranças
#blibiotecas instaladas - openpyxl - pillow (pip instal # no terminal para instalar)

import openpyxl
import webbrowser
import pyautogui
from urllib.parse import quote
from time import sleep

webbrowser.open('https://web.whatsapp.com/')
sleep(30)
#validação para aguardar o login no WhatsApp Web


#Etapa 1 Processo de leitura de dados da planilha para contatos.
workbook = openpyxl.load_workbook('Lojas.xlsx')
pagina_lojas = workbook['Lojas']

for linha in pagina_lojas.iter_rows(min_row=2):
    Loja = linha[0].value
    Nome = linha[1].value
    Numero = linha[2].value
    vencimento = linha[3].value
   #Coleta e validação dos dados da planilha.
   #Print(Loja, Numero, vencimento)

    try:
       #Mensagem personalizada para o WhatsApp.
        mensagem = f'Olá {Nome}, a ultima transmissão de dados ao SNGPC da loja ${Loja} foi em: {vencimento.strftime("%d/%m/%Y")}, se necessario, nos acione no atendimento online.'
        #Criar link para o WhatsApp.
        link_mensagem_whatsapp =f'http://api.whatsapp.com/send?phone={Numero}&text={quote(mensagem)}.'
        webbrowser.open(link_mensagem_whatsapp)
        #Atenção, você deve já estar Logado no whatsappweb pelo navegador padrão
        sleep(15)

        seta = pyautogui.locateCenterOnScreen('C:\Users\T480\Documents\Projeto_Python\BotWhatsApp\img\sent.png')
        sleep(5)
        pyautogui.click(seta[0], seta[1], duration=1)
        sleep(5)

        pyautogui.hotkey('ctrl', 'w')
        #Fechar a aba do WhatsApp Web
        sleep(5)
    except:
        print(f'Não foi possível enviar a mensagem para o contato {Nome}, verifique o número {Numero} e tente novamente.')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo_erros:
            arquivo_erros.write(f'{Loja},{Nome},{Numero},{vencimento}\n')   