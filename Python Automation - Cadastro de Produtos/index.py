
""" 
   * @Python automation - case: Cadastro de Produtos dentro de um sistema
"""

import pyautogui # Biblioteca para automatizar mouse, teclado e tela
import time

# pyautogui.click - Clicar com o mouse
# pyautogui.write - Escrever texto
# pyautogui.press - Pressionar tecla
# pyautogui..hotkey - Fazer comando de teclas
# pyautogui.scrool - Rolar tela

# Pausa para não travar
pyautogui.PAUSE = 0.5

# 1 - Entrar no sistema

# 1.1 - Abrir navegador
pyautogui.press('Win') # Abrir buscador do Windows
pyautogui.write('Opera') # Buscar Navegador Ex: Opera
pyautogui.press('Enter') # Enter para abrir navegador

# 1.2 - Entrar no no sistema
pyautogui.write('url_da_pagina_de_login_do_sistema') # Escrever caminho do sistema
pyautogui.press('Enter') # Entrar para enviar url e entrar na página

time.sleep(3)

# 2 - Fazer login no sistema
pyautogui.click(x=830, y=464)
pyautogui.hotkey('ctrl', 'a')
pyautogui.write('seu_email@gmail.com')
pyautogui.press('Tab') # passar para compo de senha
pyautogui.write('senha') # Inserir senha alatoria
pyautogui.click(x=941, y=669)

time.sleep(3)

# 3 - Importar base de dados

import pandas
tabela = pandas.read_csv('produtos.csv')

# 4 - Cadastrar produtos para cada item da tabela

for linha in tabela.index:  
   # Código
   pyautogui.click(x=780, y=340)
   codigo = str(tabela.loc[linha, 'codigo'])
   pyautogui.write(codigo)
   # Marca
   pyautogui.press('Tab')
   marca = str(tabela.loc[linha, 'marca'])
   pyautogui.write(marca)
   # Tipo
   pyautogui.press('Tab')
   tipo = str(tabela.loc[linha, 'tipo'])
   pyautogui.write(tipo)
   # Categoria
   pyautogui.press('Tab')
   categoria = str(tabela.loc[linha, 'categoria'])
   pyautogui.write(categoria)
   # Preço
   pyautogui.press('Tab')
   preco = str(tabela.loc[linha, 'preco_unitario'])
   pyautogui.write(preco)
   # Custo
   pyautogui.press('Tab')
   custo = str(tabela.loc[linha, 'custo'])
   pyautogui.write(custo)
   # Bbs
   pyautogui.press('Tab')
   obs = str(tabela.loc[linha, 'obs'])
   
   if obs != 'nan':
      pyautogui.write(obs)
      
   # Botão de enviar
   pyautogui.press('Tab')
   pyautogui.press('Enter')

   # Voltar ao inicio
   pyautogui.scroll(5000)
