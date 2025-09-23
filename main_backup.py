import pyautogui
import time
import keyboard
import pandas as pd

# Dicionário de variáveis globais, incluindo planilhas
variaveis = {}

# ========== Ações ==========
def clicar_imagem(nome):
    caminho = f'imagens/{nome.strip()}'
    print(f'🔍 Procurando {caminho}...')
    pos = pyautogui.locateCenterOnScreen(caminho, confidence=0.8)
    if pos:
        pyautogui.click(pos)
        print(f'✅ Cliquei em {nome}')
    else:
        print(f'❌ Imagem {nome} não encontrada.')

def clicar_offset(nome, offset_x, offset_y):
    caminho = f'imagens/{nome.strip()}'
    print(f'🔍 Procurando {caminho} com offset...')
    pos = pyautogui.locateCenterOnScreen(caminho, confidence=0.8)
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
        # Extrai o nome da coluna
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

def aguardar_imagem(nome, timeout=30):
    caminho = f"imagens/{nome.strip()}"
    print(f"⏳ Aguardando imagem {nome} para continuar...")
    inicio = time.time()

    while time.time() - inicio < timeout:
        print(time.time() - inicio)
        pos = None
        try:
            pos = pyautogui.locateCenterOnScreen(caminho, confidence=0.8)
        except Exception as e:
            print(f"⚠️ Falha ao tentar localizar imagem: {nome}")

        if pos:
            print(f"✅ Imagem {nome} encontrada. Continuando...")
            return True
        else:
            time.sleep(0.5)

    print(f"⚠️ Imagem {nome} não encontrada após {timeout} segundos.")
    return False


# ========== Executor ==========
def interpretar_receita(caminho='receita.txt'):
    with open(caminho, 'r', encoding='utf-8') as f:
        linhas = f.readlines()

    i = 0
    while i < len(linhas):
        linha = linhas[i].strip()

        if not linha or linha.startswith('#'):
            i += 1
            continue

        # --- Leitura de planilha ---
        if linha.startswith('ler_planilha:'):
            partes = linha.split(':')[1].strip().split()
            nome_arquivo = partes[0]
            nome_variavel = partes[1]
            ler_planilha(nome_arquivo, nome_variavel)

        # --- Início de laço de repetição ---
        elif linha.startswith('inicio_iteracao'):
            nome_df = linha.split()[1]
            df = variaveis[nome_df]
            bloco = []

            # Coleta linhas até o fim do bloco
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
        


# ========== Execução de cada linha ==========
def executar_linha(linha, linha_atual=None):
    if not linha or linha.startswith('#'):
        return

    if linha.startswith('clicar_imagem:'):
        nome = linha.split(':', 1)[1]
        clicar_imagem(nome)

    elif linha.startswith('clicar_offset:'):
        partes = linha.split(':', 1)[1].split(',')
        clicar_offset(*partes)

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
            x, y = partes
            capturar_imagem(x.strip(), y.strip())
        if len(partes) == 3:
            x, y, nome = partes
            capturar_imagem(x.strip(), y.strip(), nome.strip())
            
    elif linha.startswith('aguardar_imagem:'):
        nome = linha.split(':', 1)[1]
        aguardar_imagem(nome.strip())  

    else:
        print(f'❓ Comando desconhecido: {linha}')

# ========== Início ==========
if __name__ == '__main__':
    print("🟡 Posicione o sistema. Início em 5 segundos...")
    time.sleep(5)
    interpretar_receita()