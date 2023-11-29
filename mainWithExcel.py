import tkinter as tk
from tkinter import simpledialog
import pyautogui
import time
import random
import pandas as pd

# Variáveis para armazenar os tempos mínimos e máximos
min_value_global = 0
max_value_global = 0

def localizar_e_clicar(imagem_path, confianca=0.8):
    try:
        x, y, largura, altura = pyautogui.locateOnScreen(imagem_path, confidence=confianca)
        ponto_central_x = x + largura // 2
        ponto_central_y = y + altura // 2
        pyautogui.click(ponto_central_x, ponto_central_y)
        return True
    except TypeError:
        print("Elemento não encontrado.")
        return False

def sleep_random():
    # Aguardar um tempo aleatório dentro do intervalo definido globalmente
    sleep_time = random.uniform(min_value_global, max_value_global)
    time.sleep(sleep_time)

def obter_valores():
    global min_value_global
    global max_value_global

    # Abrir diálogo para obter valores
    min_value_global = simpledialog.askfloat("Tempo Mínimo", "Digite o valor mínimo:")
    max_value_global = simpledialog.askfloat("Tempo Máximo", "Digite o valor máximo:")

    # Aguardar um tempo aleatório dentro do intervalo
    sleep_random()

def ler_excel_colar_enter(caminho_arquivo, linhas_escolhidas, linha_anterior):
    # Ler o arquivo Excel
    df = pd.read_excel(caminho_arquivo)

    # Filtrar as linhas que ainda não foram escolhidas
    linhas_disponiveis = list(set(range(len(df))) - linhas_escolhidas)

    # Verificar se há pelo menos duas linhas disponíveis
    if len(linhas_disponiveis) >= 2:
        # Escolher uma linha aleatória que ainda não foi escolhida
        linha_coluna1 = random.choice(linhas_disponiveis)
        
        # Evitar escolher a mesma linha consecutivamente com o mesmo valor
        while linha_coluna1 == linha_anterior or df.iloc[linha_coluna1, 0] == df.iloc[linha_anterior, 0]:
            linha_coluna1 = random.choice(linhas_disponiveis)
        
        valor_coluna1 = df.iloc[linha_coluna1, 0]

        # Remover a linha escolhida das linhas disponíveis
        linhas_disponiveis.remove(linha_coluna1)

        linha_coluna2 = random.choice(linhas_disponiveis)
        
        # Evitar escolher a mesma linha consecutivamente com o mesmo valor
        while linha_coluna2 == linha_coluna1 or df.iloc[linha_coluna2, 1] == df.iloc[linha_coluna1, 0]:
            linha_coluna2 = random.choice(linhas_disponiveis)
        
        valor_coluna2 = df.iloc[linha_coluna2, 1]

        # Adicionar as linhas escolhidas ao conjunto
        linhas_escolhidas.add(linha_coluna1)
        linhas_escolhidas.add(linha_coluna2)

        # Substituir os valores nas funções pyautogui.write
        pyautogui.write(valor_coluna1, interval=0.1)  # Ajuste o intervalo conforme necessário
        sleep_random()  # Aguardar um tempo aleatório dentro do intervalo
        pyautogui.hotkey('Enter')
        sleep_random()  # Aguardar um tempo aleatório dentro do intervalo

        pyautogui.write(valor_coluna2, interval=0.1)  # Ajuste o intervalo conforme necessário
        sleep_random()  # Aguardar um tempo aleatório dentro do intervalo
        pyautogui.hotkey('Enter')

        # Retornar a linha atual para ser utilizada na próxima iteração
        return linha_coluna1
    else:
        print("Não há pelo menos duas linhas disponíveis para escolha.")

# Exemplo de uso
filtrar_mensagem = 'C:\\Users\\matheus_rsilva\\Documents\\Lightshot\\barra_pesquisa.png'

# Conjunto para armazenar as linhas já escolhidas
linhas_escolhidas = set()

# Flag para verificar se é a primeira iteração
primeira_iteracao = True

# Índice para rastrear a próxima linha a ser lida
linha_anterior = -1

# Loop para repetir o processo
while True:
    # Verificar se é a primeira iteração para solicitar os tempos
    if primeira_iteracao:
        obter_valores()
        primeira_iteracao = False

    elemento_encontrado = localizar_e_clicar(filtrar_mensagem)

    # Aguardar um tempo aleatório dentro do intervalo
    sleep_random()

    # Continuar com o restante do seu script
    linha_anterior = ler_excel_colar_enter('C:\\Users\\matheus_rsilva\\Desktop\\ChatBot_WhatsApp_PY\\mensagens.xlsx', linhas_escolhidas, linha_anterior)

    sleep_random()
    pyautogui.hotkey('esc')
    sleep_random()

    # Posicionar o mouse no centro da tela após a execução da primeira vez
    if not linhas_escolhidas:
        pyautogui.moveTo(pyautogui.size()[0] // 2, pyautogui.size()[1] // 2)

    # Adicionar um tempo de espera específico antes de repetir o processo
    sleep_random()
