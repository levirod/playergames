import pandas as pd

# Carregar a planilha
df = pd.read_excel(r'C:\Users\venda\OneDrive\Área de Trabalho\Resolver Planilha\planilha_custo_semlojas.xlsx')

# Agrupar os produtos pelo nome e obter o maior valor de custo para cada grupo
df['Custo'] = df.groupby('Descrição do Produto')['Custo'].transform('max')

# Salvar a planilha atualizada
df.to_excel(r'C:\Users\venda\OneDrive\Área de Trabalho\Resolver Planilha\planilha_custo_atualizada.xlsx', index=False)
