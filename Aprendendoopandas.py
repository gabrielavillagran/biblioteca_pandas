import pandas as pd
import win32com.client as win32

# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
'''pd.set_option('display.max_columns', None)'''

'''print(tabela_vendas)'''

# Faturamento por loja

"""filtrando as colunas na tabela"""
filtro1 = tabela_vendas[['ID Loja', 'Valor Final']]

"""agrupando as lojas vamos somar o valor final """
faturamento = filtro1.groupby('ID Loja').sum()
faturamento = faturamento.sort_values(by='Valor Final', ascending=False)

print(faturamento)
print('-' * 50)

# Quantidade de produtos vendidos por loja

filtro2 = tabela_vendas[['ID Loja', 'Quantidade']]

quantidade = filtro2.groupby('ID Loja').sum()
quantidade = quantidade.sort_values(by='Quantidade', ascending=False)

print(quantidade)
print('-' * 50)

# Ticket médio por produto em cada loja

"""Preciso dividir as COLUNAS de faturamento/quantidade"""

"""Sempre que dividimos uma coluna pela outra ou multiplicamos, 
e queremos que seja TABELA colocamos to_frame() """

ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
ticket_medio = ticket_medio.sort_values(by='Ticket Médio', ascending=False)
print(ticket_medio)
print('-' * 50)

# Enviar um e-mail com o relatório
"""importamos a biblioteca pywin32, buscar no google """
# <p> Texto para Parágrafo </p>

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'villagranfurg@gmail.com'
mail.Subject = 'Relatório de vendas do Excel para Python'
mail.HTMLBody = f'''
<p>Esse é o relatório de vendas que foi feito atráves do Python 
transportando a tabela do excel!!</p>


<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{quantidade.to_html()}


<p>Ticket médio dos produtos em cada loja</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p> Medidas estatísticas:</p>
{tabela_vendas.describe().round(2).to_html()}

<p>Att. Gabriela</p>

'''  # this field is optional
# to_html() ele formata como tabela no codigo HTML
# To attach a file to the email (optional):
# Formatter, para formatar em R$ as colunas dos números.

mail.Send()
print('email enviado')
