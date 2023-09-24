import pandas as pd
import win32com.client as win32

# Importar base de dados
tabela_vendas = pd.read_excel('Vendas (1).xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)

# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)

# Ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# Enviar email com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'joao.alves6@outlook.com;'  #após o ; adiciona-se um novo email
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''

<p>Olá,</p>
<p>segue abaixo as tabelas de faturamento das lojas</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade de vendas:</p>
{quantidade.to_html()}

<p>Ticket médio dos produtos por loja</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}



<p>Qualquer dúvida, estamos à disposição. Basta responder a este e-mail.</p>
<p>Att.</p>
<p>João</p>
'''

# Enviar o email
mail.Send()
print('E-mail enviado com sucesso!')
