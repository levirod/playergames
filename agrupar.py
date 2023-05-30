import pandas as pd

# Carregar os dados da primeira planilha
planilha1 = pd.read_excel(r'C:\Users\venda\OneDrive\Área de Trabalho\Resolver Planilha\planilha1.xlsx')

# Carregar os dados da segunda planilha
planilha2 = pd.read_excel(r'C:\Users\venda\OneDrive\Área de Trabalho\Resolver Planilha\planilha2.xlsx')

# Selecionar as colunas relevantes da planilha2
planilha2 = planilha2[['Código NCM', 'CFOP', 'CEST']]

# Realizar o merge das planilhas com base no Código NCM
planilha_atualizada = pd.merge(planilha1, planilha2, on='Código NCM', how='left')

# Atualizar os valores de CFOP e CEST na planilha1 com os valores correspondentes
# da planilha2
planilha_atualizada['CFOP_x'] = planilha_atualizada['CFOP_y']
planilha_atualizada['CEST_x'] = planilha_atualizada['CEST_y']

# Descartar as colunas desnecessárias
planilha_atualizada = planilha_atualizada.drop(['CFOP_y', 'CEST_y'], axis=1)

# Salvar a planilha atualizada em um novo arquivo
planilha_atualizada.to_excel(r'C:\Users\venda\OneDrive\Área de Trabalho\Resolver Planilha\planilha1_atualizada.xlsx', index=False)

#modificação nova23
