import openpyxl as px
import webbrowser
from urllib.parse import quote
from time import sleep
import pyautogui

pyautogui.FAILSAFE = True  
# basicamente você aborta o programa se mover o mouse bem pro canto

# 1. Abre o WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
sleep(15)  
# Isso aqui da um tempo maior quando voçê vai abrir pela primeira vez, por que muitas vezes você vai fazer abrir porém ele não ta logado ainda

# 2. Abre a planilha
workbook = px.load_workbook('Bot_WhatsApp.xlsx')
pagina_clientes = workbook['Página1']

for linha in pagina_clientes.iter_rows(min_row=2, values_only=True):
    nome, telefone = linha[0], linha[1]

    if not nome or not telefone:
        continue  # ignora linhas inválidas

    mensagem = (
        f"Olá {nome}, tudo bem? "
        f"Esta é uma mensagem de teste feita por um bot de WhatsApp. "
        f"Apenas ignore ela."
    )

    link = (
        "https://web.whatsapp.com/send"
        f"?phone={telefone}&text={quote(mensagem)}"
    )

    # 3. Abre conversa
    webbrowser.open(link)
    sleep(12)

    # 4. Procura o botão de enviar (com tentativas)
    botao = None
    for _ in range(10):  # tenta por 10 segundos
        botao = pyautogui.locateCenterOnScreen('assets/botao.png',confidence=0.8)
        if botao:
            break
        sleep(1)

    if not botao:
        print(f"Error: Botão não encontrado para {telefone}")
        pyautogui.hotkey('ctrl', 'w')
        sleep(2)
        continue

    # 5. Clica no botão de enviar
    pyautogui.click(botao)
    sleep(2)

    # 6. Fecha a aba
    pyautogui.hotkey('ctrl', 'w')
    sleep(3)
