# automacao_de_tarefa
Código para automação e alimentação de planilha em Python


import pyautogui
import time
import pyperclip
from openpyxl import load_workbook

def limpar_salario(valor):
    try:
        valor_limpo = f"{float(valor):.2f}"
    except ValueError:
        valor_limpo = "0.00"
    return valor_limpo

def arredondar_horas(tempo):
    tempo_str = str(tempo)
    if tempo_str == '0' or tempo_str == '' or ':' not in tempo_str:
        return '0'
    try:
        horas, minutos = map(int, tempo_str.split(':'))
        return str(horas + (1 if minutos > 0 else 0))
    except ValueError:
        return '0'
    
def copiar_texto(x, y):
    pyautogui.click(x, y)
    pyautogui.click(x, y, clicks=2)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.5)
    return pyperclip.paste()


# Carregar a planilha existente
workbook = load_workbook(filename='Horas Extras.xlsx')
sheet = workbook.active

# Configurar o navegador
pyautogui.PAUSE = 1
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
time.sleep(1)
pyautogui.getWindowsWithTitle("Chrome")[0].maximize()
time.sleep(1)
pyautogui.write("https://www.idinheiro.com.br/calculadoras/calculadora-horas-extras/#calculator-result")
pyautogui.press("enter")
time.sleep(3)
pyautogui.scroll(5000)

for linha in range(2, sheet.max_row + 1):  # Inicia na linha 2 para pular o cabeçalho
    if sheet[f'C{linha}'].value:  # Verifica se há um valor no campo salário base
        pyautogui.scroll(10000)
        salario_base = limpar_salario(sheet[f'C{linha}'].value)
        he_50 = arredondar_horas(sheet[f'D{linha}'].value)
        he_100 = arredondar_horas(sheet[f'E{linha}'].value)
        
        # Preenchimento dos campos
        # Campo Salário Bruto
        pyautogui.click(221, 421)  # Clica no campo Salário Bruto
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(str(salario_base))
        pyautogui.press('tab')

        # Campo Horas Trabalhadas no Mês
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write("220")
        pyautogui.press('tab')

        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write("0")
        pyautogui.press('tab')

        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(he_100)
        pyautogui.press('tab')
      
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(he_50)
        pyautogui.press('tab')

        # Continuar com os outros campos
        # Enviar o formulário e pausar para permitir carregamento da página
        pyautogui.press("enter")
        time.sleep(2)

        pyautogui.scroll(-200)
        time.sleep(1)
        
        # Capturar e processar valores
        valor_he_100 = copiar_texto(680, 662)
        valor_he_50 = copiar_texto(680, 705)
        
        sheet[f'F{linha}'].value = valor_he_50
        sheet[f'G{linha}'].value = valor_he_100
       

# Salvar a planilha após processar todas as linhas
workbook.save("Horas Extras Atualizada.xlsx")




