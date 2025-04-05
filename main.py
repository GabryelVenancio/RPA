import pandas as pd
import pyautogui
import time
import os
import random
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment

class AutomacaoG1:
    def __init__(self):
        self.relatorio = []
        self.inicio_execucao = datetime.now()
        self.criar_diretorios()
        
    def criar_diretorios(self):
        os.makedirs("screenshots", exist_ok=True)
        
    def ler_tarefas(self, arquivo):
        try:
            return pd.read_csv(arquivo)
        except Exception as e:
            print(f"Erro ao ler arquivo: {e}")
            return None

    def executar_tarefa(self, tarefa, tipo, dado):
        inicio = time.time()
        status = "Sucesso"
        
        try:
            if tipo == 'executavel':
                self.abrir_executavel(dado)
            elif tipo == 'navegar':
                self.acessar_url(dado)
            elif tipo == 'scroll':
                pyautogui.scroll(-int(dado))
            elif tipo == 'click_aleatorio':
                y_min, y_max = map(int, dado.split('-'))
                self.clicar_aleatorio(y_min, y_max)
            elif tipo == 'espera':
                time.sleep(float(dado))
            else:
                status = "Tipo de ação não reconhecido"
        except Exception as e:
            status = f"Falha: {str(e)[:50]}"
        
        tempo = round(time.time() - inicio, 2)
        self.registrar_execucao(tarefa, tipo, dado, status, tempo)
        self.capturar_tela(tarefa)

    def abrir_executavel(self, programa):
        pyautogui.hotkey('winleft', 'r')
        pyautogui.write(programa)
        pyautogui.press('enter')
        time.sleep(5)

    def acessar_url(self, url):
        pyautogui.hotkey('ctrl', 'l')
        pyautogui.write(url)
        pyautogui.press('enter')
        time.sleep(5)

    def clicar_aleatorio(self, y_min, y_max):
        """Clica em uma posição aleatória na área de notícias"""
        largura = pyautogui.size().width
        x = random.randint(int(largura*0.25), int(largura*0.75))  # Evita bordas
        y = random.randint(y_min, y_max)
        
        pyautogui.moveTo(x, y, duration=0.5)
        pyautogui.click()
        print(f"Clicado em: ({x}, {y})")

    def capturar_tela(self, nome_arquivo):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pyautogui.screenshot(f"screenshots/{nome_arquivo}_{timestamp}.png")

    def registrar_execucao(self, tarefa, tipo, dado, status, tempo):
        self.relatorio.append({
            'Tarefa': tarefa,
            'Tipo': tipo,
            'Dado': dado,
            'Status': status,
            'Tempo (s)': tempo,
            'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })

    def gerar_relatorio(self):
        df = pd.DataFrame(self.relatorio)
        
        with pd.ExcelWriter('relatorio.xlsx', engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Relatório')
            
            workbook = writer.book
            worksheet = writer.sheets['Relatório']
            
            # Formatação
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = openpyxl.styles.PatternFill(
                start_color="4F81BD", 
                end_color="4F81BD", 
                fill_type="solid"
            )
            
            for col in worksheet[1]:
                col.font = header_font
                col.fill = header_fill
                col.alignment = Alignment(horizontal='center')
            
            # Ajuste de colunas
            column_widths = {'A': 25, 'B': 15, 'C': 20, 'D': 25, 'E': 15, 'F': 20}
            for col, width in column_widths.items():
                worksheet.column_dimensions[col].width = width

def main():
    automacao = AutomacaoG1()
    tarefas = automacao.ler_tarefas('tarefas.csv')
    
    if tarefas is not None:
        for _, row in tarefas.iterrows():
            automacao.executar_tarefa(row['Tarefa'], row['Tipo'], row['Dado'])
        
        automacao.gerar_relatorio()
        print("Automação concluída! Relatório gerado.")

if __name__ == "__main__":
    main()