"""
   * Importante * - Executar arquivo no Terminal como Administrator, para evitar erro
"""

import pyautogui
import time

pyautogui.FAILSAFE = False

# Função para Entrar na Pasta e Remover todos Itens
def remove_files():
   pyautogui.press('enter') # Entrar
   pyautogui.hotkey('ctrl', 'a') # Selecionar Todos Arquivos
   pyautogui.hotkey('shift', 'delete') # Deletar
   pyautogui.press('enter') # Confirmar Ação

pyautogui.PAUSE = 0.5
comand_list = ['prefetch', 'recent', 'temp']

for comand in comand_list:
   pyautogui.hotkey('win', 'r')
   pyautogui.write(comand)
   remove_files()
   pyautogui.hotkey('alt', 'f4')
   
time.sleep(3)
path_temp = r'C:\Users\Mr. Specter\AppData\Local\Temp'
pyautogui.hotkey('win', 'e')

time.sleep(3)
pyautogui.click(x=414, y=64)
pyautogui.write(path_temp)

time.sleep(3)
remove_files()

# Fechar Janelas do Explorador de Arquivos
for i in range(4):
   pyautogui.hotkey('alt', 'f4'), time.sleep(1.5)
