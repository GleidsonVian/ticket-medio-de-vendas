import pandas as pd
import win32com.client as win32

# Lê o arquivo Excel 'Vendas.xlsx' e armazena na variável tabela_vendas
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Agrupa os dados por 'ID Loja' e calcula a soma do 'Valor Final' para cada loja
vendas = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

# Agrupa os dados por 'ID Loja' e calcula a soma da 'Quantidade' para cada loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# Calcula o ticket médio (Valor Final dividido pela Quantidade) e converte para DataFrame
ticket_medio = (vendas['Valor Final'] / quantidade['Quantidade']).to_frame()

# Renomeia a coluna do DataFrame resultante para 'Ticket Médio'
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})

# Cria uma instância do Outlook
outlook = win32.Dispatch('outlook.application')

# Cria um novo item de email
mail = outlook.CreateItem(0)

# Define o destinatário do email
mail.To = 'gleidson.testes1@outlook.com'

# Define o assunto do email
mail.Subject = 'Relatorio de Vendas por Loja'

# Define o corpo do email em HTML, incluindo as tabelas de vendas, quantidade e ticket médio
mail.HTMLBody = f"""
<p>Prezados colaboradores</p>
<p>Segue o Relatorio de Vendas por cada Loja</p>

<p>Faturamento:
{vendas.to_html()}
 

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html()}
"""

# Envia o email
mail.Send()

# Imprime mensagem de confirmação no console
print('Email enviado!')
