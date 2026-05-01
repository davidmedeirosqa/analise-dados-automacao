# Importar a base de dados
import pandas as pd
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# Faturamento da loja

# Quantidade de produtos vendidos por loja

# Ticket médio por produto em cada loja

# Enviar um e-mail com o relatório