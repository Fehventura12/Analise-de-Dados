import win32com.client as win32

# Importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

# Ticket médio por produto em cada loja
tkm = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(tkm)
tkm = tkm.rename(colunms={0: 'Ticket Médio'}) # Renomenado o nome da coluna
# comando to_frame() serve para deixar os dados como tabela

# Enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'fernando_sventura@outlook.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:<p/>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html}

<p>Tickt Médio dos Produtos em cada Loja.</p>
{tkm.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou a disposição.</p>

<p>Att...</p>
<p>Fernando</p>
'''

mail.Send()

# Esse é um relatório simples enviado automáticamente por Email