import pandas as pd
import win32com.client as win32
import matplotlib.pyplot as plt

# Importar a base de dados
tabela = pd.read_excel('Vendas_bermudas.xlsx')


# Visualizar a base de dados
"""pd.set_option('display.max_columns', None)

print(tabela)"""

# Faturamento por loja
# .sort_values(by=' ', ascending=False) coloca em ordem crescente ou decrescente

filtro1 = tabela[['ID Loja', 'Valor Final']]

faturamento = filtro1.groupby('ID Loja').sum()
faturamento = faturamento.sort_values(by='Valor Final', ascending=False)

print(faturamento)
print('-' * 50)

fig, ax = plt.subplots()
plt.hist(filtro1['Valor Final'], bins=100, color='turquoise')
ax.set_title('Valor Final por Lojas')
ax.set_ylabel('Valor Final')
# substituir o nome do eixo x
ax.set_xticklabels(filtro1['ID Loja'])
'''
def autolabel(filtro1):
    for i in filtro1:
        h = i.get_height()
        ax.annotate('{}'.format(h),
                xy = (i.get_x()+i.get_width()/2,h),
                xytext = (0,3),
                textcoords = 'offset points',
                ha ='center')

plt.show()
'''


# Quantidade de produtos vendidos por loja

filtro2 = tabela[['ID Loja', 'Quantidade']]

quantidade = filtro2.groupby('ID Loja').sum()
quantidade = quantidade.sort_values(by='Quantidade', ascending=False)
print(quantidade)
print('-'*50)

# Ticket médio por produto em cada loja

ticket_medio = (faturamento['Valor Final']/quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
ticket_medio = ticket_medio.sort_values(by='Ticket Médio', ascending=False)

print(ticket_medio)
print('-'*50)

# Media, Mediana, Moda, são consideradas apenas colunas com formato numérico

print(tabela.describe())

# Enviar um e-mail com o relatório
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
{tabela.describe().round(2).to_html()}

<p>Att. Gabriela</p>

'''

mail.Send()
print('email enviado')
