import pandas as pd

# Carregar os dados da primeira planilha
planilha1 = pd.read_excel(r'C:\Users\venda\OneDrive\Área de Trabalho\Resolver Planilha\planilha1.xlsx')

# Carregar os dados da segunda planilha com os códigos NCM, CFOP e CEST atualizados
planilha2 = pd.read_excel(r'C:\Users\venda\OneDrive\Área de Trabalho\Resolver Planilha\planilha2.xlsx')

# Atualizar os valores de CFOP e CEST com base nos códigos NCM da planilha2
for index, row in planilha2.iterrows():
    codigo_ncm = row['Código NCM']
    cfop = row['CFOP']
    cest = row['CEST']
    planilha1.loc[planilha1['Código NCM'] == codigo_ncm, 'CFOP'] = cfop
    planilha1.loc[planilha1['Código NCM'] == codigo_ncm, 'CEST'] = cest

# Salvar a planilha atualizada em um novo arquivo
planilha1.to_excel(r'C:\Users\venda\OneDrive\Área de Trabalho\Resolver Planilha\planilha1_atualizada.xlsx', index=False)
