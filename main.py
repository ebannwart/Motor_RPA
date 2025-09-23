import pyautogui
import time
import keyboard
import pandas as pd
import os
import sys
from datetime import datetime

# Garante que a pasta 'log' exista
os.makedirs("log", exist_ok=True)

# Define nome do arquivo com timestamp
agora = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
log_path = os.path.join("log", f"log_{agora}.txt")

# Redireciona todos os prints para o arquivo de log
sys.stdout = open(log_path, 'w', encoding='utf-8')
sys.stderr = sys.stdout  # Também redireciona erros

# Dicionário de variáveis globais, incluindo planilhas
variaveis = {}

# ========== AÇÕES ==========
def clicar_imagem(nome, confidence=0.8):
    caminho = f'imagens/{nome.strip()}'
    print(f'🔍 Procurando {caminho}... (confidence={confidence})')
    pos = pyautogui.locateCenterOnScreen(caminho, confidence=confidence)
    if pos:
        pyautogui.click(pos)
        print(f'✅ Cliquei em {nome}')
    else:
        print(f'❌ Imagem {nome} não encontrada.')

def clicar_offset(nome, offset_x, offset_y, confidence=0.8):
    caminho = f'imagens/{nome.strip()}'
    print(f'🔍 Procurando {caminho} com offset... (confidence={confidence})')
    pos = pyautogui.locateCenterOnScreen(caminho, confidence=confidence)
    if pos:
        x, y = pos
        pyautogui.click(x + int(offset_x), y + int(offset_y))
        print(f'✅ Cliquei com offset em {nome}')
    else:
        print(f'❌ Imagem {nome} não encontrada.')

def clicar_posicao(x, y):
    x = int(x)
    y = int(y)
    pyautogui.click(x, y)
    print(f'🖱️ Cliquei na posição ({x}, {y})')

def escrever(texto, linha_atual=None):
    if texto.strip().startswith("plan["):
        col = texto.strip().replace("plan['", "").replace("']", "")
        valor = str(linha_atual[col]) if linha_atual is not None else ""
        pyautogui.write(valor, interval=0.05)
        print(f'📝 Escrevi valor da planilha: {valor}')
    else:
        pyautogui.write(texto.strip(), interval=0.05)
        print(f'📝 Escrevi texto fixo: {texto.strip()}')

def esperar(segundos):
    print(f'⏳ Esperando {segundos} segundos...')
    time.sleep(float(segundos))

def esperar_imagem(nome, timeout=15):
    caminho = f'imagens/{nome.strip()}'
    print(f'⏳ Esperando imagem {nome} aparecer...')
    inicio = time.time()
    while time.time() - inicio < timeout:
        pos = pyautogui.locateCenterOnScreen(caminho, confidence=0.8)
        if pos:
            print(f'✅ Imagem {nome} encontrada!')
            return
        time.sleep(0.5)
    print(f'❌ Imagem {nome} não apareceu em {timeout}s.')

def pressionar_tecla(tecla):
    pyautogui.press(tecla.strip())
    print(f'⌨️ Pressionei tecla: {tecla}')

def combinar_teclas(*teclas):
    pyautogui.hotkey(*[t.strip() for t in teclas])
    print(f'🎹 Comando de teclas combinado: {", ".join(teclas)}')

def ler_planilha(nome_arquivo, nome_variavel):
    df = pd.read_excel(nome_arquivo)
    variaveis[nome_variavel] = df
    print(f'📊 Planilha {nome_arquivo} carregada como "{nome_variavel}" com {len(df)} linhas')

def capturar_imagem(x, y, nome_arquivo=None):
    largura = 200
    altura = 200
    x = int(x)
    y = int(y)
    topo_esquerdo_x = x - largura // 2
    topo_esquerdo_y = y - altura // 2
    if not nome_arquivo:
        timestamp = int(time.time())
        nome_arquivo = f"captura_{x}_{y}_{timestamp}.png"
    caminho = f"imagens/{nome_arquivo}"
    img = pyautogui.screenshot(region=(topo_esquerdo_x, topo_esquerdo_y, largura, altura))
    img.save(caminho)
    print(f"📸 Imagem salva em: {caminho}")

def aguardar_imagem(nome, timeout=30, confidence=0.8):
    caminho = f"imagens/{nome.strip()}"
    print(f"⏳ Aguardando imagem {nome}... (confidence={confidence})")
    inicio = time.time()
    while time.time() - inicio < timeout:
        try:
            pos = pyautogui.locateCenterOnScreen(caminho, confidence=confidence)
            if pos:
                print(f"✅ Imagem {nome} encontrada. Continuando...")
                return True
        except Exception as e:
            print(f"⚠️ Erro ao localizar imagem {nome}: {e}")
            return False
        time.sleep(0.5)
    print(f"❌ Imagem {nome} não encontrada após {timeout} segundos.")
    return False


def clicar_e_arrastar_offset(nome, offset_x, offset_y, duracao=1.0, confidence=0.8):
    caminho = f'imagens/{nome.strip()}'
    print(f'🔍 Procurando {caminho} para arrastar... (confidence={confidence})')
    pos = pyautogui.locateCenterOnScreen(caminho, confidence=confidence)
    if pos:
        x, y = pos
        destino_x = x + int(offset_x)
        destino_y = y + int(offset_y)
        pyautogui.moveTo(x, y)
        pyautogui.mouseDown()
        pyautogui.moveTo(destino_x, destino_y, duration=float(duracao))
        pyautogui.mouseUp()
        print(f'✅ Arrastei de ({x}, {y}) para ({destino_x}, {destino_y}) em {duracao} segundos.')
    else:
        print(f'❌ Imagem {nome} não encontrada para arrastar.')

def duplo_clique_imagem(nome, confidence=0.8):
    caminho = f'imagens/{nome.strip()}'
    print(f'🔍 Procurando {caminho} para duplo clique... (confidence={confidence})')
    pos = pyautogui.locateCenterOnScreen(caminho, confidence=confidence)
    if pos:
        pyautogui.doubleClick(pos)
        print(f'✅ Duplo clique realizado em {nome}')
    else:
        print(f'❌ Imagem {nome} não encontrada.')
        
def rolar_scroll(quantidade):
    quantidade = int(quantidade)
    pyautogui.scroll(quantidade)
    direcao = 'cima' if quantidade > 0 else 'baixo'
    print(f'🖱️ Scroll para {direcao} ({quantidade} unidades)')


# ========== EXECUTOR ==========
def interpretar_receita(caminho='receita.txt'):
    with open(caminho, 'r', encoding='utf-8') as f:
        linhas = f.readlines()

    i = 0
    while i < len(linhas):
        linha = linhas[i].strip()
        if not linha or linha.startswith('#'):
            i += 1
            continue

        if linha.startswith('ler_planilha:'):
            partes = linha.split(':')[1].strip().split()
            ler_planilha(partes[0], partes[1])

        elif linha.startswith('inicio_iteracao'):
            nome_df = linha.split()[1]
            df = variaveis[nome_df]
            bloco = []
            i += 1
            while i < len(linhas) and not linhas[i].strip().startswith(f'fim_iteracao {nome_df}'):
                bloco.append(linhas[i].strip())
                i += 1
            for idx, linha_atual in df.iterrows():
                print(f'🔁 Iteração {idx + 1}/{len(df)}')
                for comando in bloco:
                    executar_linha(comando, linha_atual)
        else:
            executar_linha(linha)
        i += 1

# ========== EXECUTA UMA LINHA ==========
def executar_linha(linha, linha_atual=None):
    if not linha or linha.startswith('#'):
        return

    if linha.startswith('clicar_imagem:'):
        partes = linha.split(':', 1)[1].split(',')
        if len(partes) == 2:
            clicar_imagem(partes[0].strip(), float(partes[1].strip()))
        else:
            clicar_imagem(partes[0].strip())

    elif linha.startswith('clicar_offset:'):
        partes = linha.split(':', 1)[1].split(',')
        if len(partes) == 4:
            clicar_offset(partes[0].strip(), partes[1], partes[2], float(partes[3]))
        else:
            clicar_offset(partes[0].strip(), partes[1], partes[2])

    elif linha.startswith('clicar_posicao:'):
        x, y = linha.split(':', 1)[1].split(',')
        clicar_posicao(x.strip(), y.strip())

    elif linha.startswith('escrever:'):
        texto = linha.split(':', 1)[1]
        escrever(texto, linha_atual)

    elif linha.startswith('esperar_imagem:'):
        nome = linha.split(':', 1)[1]
        esperar_imagem(nome)

    elif linha.startswith('esperar:'):
        segundos = linha.split(':', 1)[1]
        esperar(segundos)

    elif linha.startswith('pressionar_tecla:'):
        tecla = linha.split(':', 1)[1]
        pressionar_tecla(tecla)

    elif linha.startswith('combinar_teclas:'):
        teclas = linha.split(':', 1)[1].split(',')
        combinar_teclas(*teclas)

    elif linha.startswith('capturar_imagem:'):
        partes = linha.split(':', 1)[1].split(',')
        if len(partes) == 2:
            capturar_imagem(partes[0].strip(), partes[1].strip())
        elif len(partes) == 3:
            capturar_imagem(partes[0].strip(), partes[1].strip(), partes[2].strip())

    elif linha.startswith('aguardar_imagem:'):
        partes = linha.split(':', 1)[1].split(',')
        if len(partes) == 2:
            aguardar_imagem(partes[0].strip(), confidence=float(partes[1].strip()))
        else:
            aguardar_imagem(partes[0].strip())
            
    elif linha.startswith('clicar_e_arrastar:'):
        partes = linha.split(':', 1)[1].split(',')
        if len(partes) == 4:
            clicar_e_arrastar_offset(partes[0].strip(), partes[1], partes[2], partes[3])
        elif len(partes) == 5:
            clicar_e_arrastar_offset(partes[0].strip(), partes[1], partes[2], partes[3], float(partes[4]))
        else:
            print(f'⚠️ Uso incorreto do comando clicar_e_arrastar: {linha}')
            
    elif linha.startswith('duplo_clique_imagem:'):
        partes = linha.split(':', 1)[1].split(',')
        if len(partes) == 2:
            duplo_clique_imagem(partes[0].strip(), float(partes[1].strip()))
        else:
            duplo_clique_imagem(partes[0].strip())    

    elif linha.startswith('scroll:'):
        quantidade = linha.split(':', 1)[1]
        rolar_scroll(quantidade)


    else:
        print(f'❓ Comando desconhecido: {linha}')

# ========== INÍCIO ==========
if __name__ == '__main__':
    print("🟡 Posicione o sistema. Início em 5 segundos...")
    time.sleep(5)
    interpretar_receita()