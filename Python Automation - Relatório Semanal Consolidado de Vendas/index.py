""" 
    # Requisitos:
    
    * (1) - Ter Instalado o Outlook
    * (2) - Ter Instalado todas Bibliotecas (schedule, pandas, win32com)
    
    # Passo a passo:

    * (1) - Importar Bibliotecas.
    * (2) - Extrair, Consolidar e Tratar as Bases Existentes.
    * (3) - Criar Email e Anexar a Planilha. 
    * (4) - Garantir Atualização Semanal p/ toda segunda-feira às 22:11 das bases existentes da pasta. Exemplo: 'base'.

"""


# Passo (1)
import os, schedule, time
import pandas as pd
import win32com.client as win32
from datetime import datetime

def atualizacao_planilha():
    # Passo (2)
    caminho = 'bases'
    files = os.listdir(caminho)
    tabela_consolidada = pd.DataFrame()

    for file_name in files:
        tabela = pd.read_csv(os.path.join(caminho, file_name))
        tabela['Data de Venda'] = pd.to_datetime('01/01/1900') + pd.to_timedelta(tabela['Data de Venda'], unit='d')
        tabela_consolidada = pd.concat([tabela_consolidada, tabela])
        
    tabela_consolidada = tabela_consolidada.sort_values(by = 'Data de Venda')
    tabela_consolidada = tabela_consolidada.reset_index(drop = True)
    tabela_consolidada.to_excel('Vendas.xlsx', index = False)

    # Passo (3)
    outlook = win32.Dispatch('outlook.application')
    
    email = outlook.CreateItem(0)
    email.To = 'seu_email@gmail.com'
    data_atual = datetime.today().strftime('%d/%m/%y')
    email.Subject = f'Relatório de Vendas {data_atual}'
    
    email.Body = f""" 

        Prezados, \n

        Em anexo, segue o relatório de vendas referente ao dia {data_atual}.

        Estou à disposição para qualquer dúvida ou esclarecimento.

        Atenciosamente,

        'Seu nome'

    """
    
    caminho = os.getcwd()
    anexo = os.path.join(caminho, 'Vendas.xlsx')
    email.Attachments.Add(anexo)
    email.Send()

# Passo (4)
schedule.every().monday.at("22:12").do(atualizacao_planilha)

while True:
    schedule.run_pending()
    time.sleep(1)
