import pandas as pd

# Ler a planilha
planilha = pd.read_excel('planilha_tagplus.xlsx')

# Filtrar as linhas que contêm "ASSISTÊNCIA TÉCNICA -" ou "ASSISTÊNCIA CONSOLE -"
planilha_filtrada = planilha[~planilha['Descrição do Produto'].str.contains('ASSISTÊNCIA TÉCNICA -|ASSISTENCIA CONSOLE -')]

# Salvar a planilha filtrada em um novo arquivo
planilha_filtrada.to_excel('planilha_tagplus.xlsx', index=False)
