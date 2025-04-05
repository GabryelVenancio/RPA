import pyautogui
import time
import random
import os
from datetime import datetime

pyautogui.PAUSE = 2
LARGURA, ALTURA = pyautogui.size()
SCROLLS_INICIAIS = 2

def configurar_ambiente():
    """Prepara pastas e ajustes iniciais"""
    os.makedirs("screenshots", exist_ok=True)
    pyautogui.FAILSAFE = True

def mostrar_progresso(passo, acao):
    """Registra e documenta cada passo"""
    print(f"\n{'='*50}")
    print(f"▶ PASSO {passo}: {acao}")
    print(f"{'='*50}")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    pyautogui.screenshot(f"screenshots/passo_{passo}_{timestamp}.png")

def abrir_navegador():
    """Abre o navegador e maximiza"""
    mostrar_progresso(1, "Abrindo navegador")
    pyautogui.hotkey('winleft', 'r')
    pyautogui.write('opera --start-maximized')
    pyautogui.press('enter')
    time.sleep(7)

def acessar_g1():
    """Navega para a página principal do G1"""
    mostrar_progresso(2, "Acessando G1")
    pyautogui.hotkey('ctrl', 'l')
    time.sleep(0.5)
    pyautogui.write('https://g1.globo.com/')
    pyautogui.press('enter')
    time.sleep(8)

def carregar_mais_noticias():
    """Rola a página para carregar mais notícias"""
    mostrar_progresso(3, "Carregando mais notícias")
    for _ in range(SCROLLS_INICIAIS):
        pyautogui.scroll(-500)
        time.sleep(2)

def clicar_noticia_aleatoria():
    """Seleciona e clica em uma notícia aleatória"""
    mostrar_progresso(4, "Selecionando notícia aleatória")

    x_min = LARGURA * 0.25
    x_max = LARGURA * 0.75
    y_min = 300
    y_max = ALTURA - 200

    x = random.randint(int(x_min), int(x_max))
    y = random.randint(y_min, y_max)

    pyautogui.moveTo(x, y, duration=random.uniform(0.5, 1.0))
    time.sleep(0.3)
    pyautogui.click()

    print(f"✔ Clicado em: X={x}, Y={y}")
    return (x, y)

def verificar_pagina_carregada():
    """Verifica se a página carregou completamente"""
    try:
        if pyautogui.locateOnScreen('imagens/logo_g1.png', confidence=0.7):
            return True
    except:
        pass

    time.sleep(5)
    return True

def main():
    configurar_ambiente()
    
    try:
        abrir_navegador()
        acessar_g1()

        if not verificar_pagina_carregada():
            print("⚠ Página não carregou corretamente")
            return

        carregar_mais_noticias()
        posicao_clicada = clicar_noticia_aleatoria()

        mostrar_progresso(5, f"Notícia em {posicao_clicada} aberta")
        time.sleep(5)
        
    except Exception as e:
        print(f"❌ Erro: {str(e)[:200]}")
    finally:
        mostrar_progresso(6, "Finalizando automação")

if __name__ == "__main__":
    for file in os.listdir("screenshots"):
        if file.startswith("passo_"):
            os.remove(f"screenshots/{file}")
    
    main()