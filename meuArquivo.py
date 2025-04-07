# Importar base de dados
import pandas as pd
import win32com.client as win32



# Visualizar base de dados
# pd.set_option('display.max_columns', None)      >>>>>>>>>>>>>>       Esse comando diz pro python mostrar a quantidade maxima de colunas. Não foi necessário pois o terminal já exibiu a .Qt maxima
tabela_vendas = pd.read_excel('Vendas.xlsx')
"\n"
"\n"
"\n"

# Faturamento
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

print('EXIBBINDO FATURAMENTO POR LOJA')
print(faturamento)
"\n"
print('-'*50)
"\n"

# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

print('EXIBBINDO QUANTIDADE DE PRODUTOS POR LOJA')
print(quantidade)
"\n"
print('-'*50)
"\n"

# Ticket medio em cada uma das loja > Faturamento por loja dividido pela quantidade de produtos por loja

print('EXIBBINDO TICKET MEDIO DOS PRODUTOS POR LOJA')
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(ticket_medio)
print('antes do email')


# Enviando email

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = '<email>' # substituir pelo email que receberá o relatório
mail.Subject = 'Relatório de analise de vendas - Fabio'
mail.HTMLBody = '<h2>HTML Message</h2>'

mail.Send()

print('Email enviado!')