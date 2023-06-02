import openpyxl

def manter_colunas(planilha, colunas_manter):
    colunas_remover = [coluna for coluna in range(1, planilha.max_column + 1) if coluna not in colunas_manter]
    for coluna in colunas_remover[::-1]:
        planilha.delete_cols(coluna)

# Abrir o arquivo do Excel
arquivo_excel = r'C:\Users\venda\OneDrive\Área de Trabalho\Resolver Planilha\estoque.xlsx'
workbook = openpyxl.load_workbook(arquivo_excel)

# Selecionar a planilha
nome_planilha = 'Worksheet'

try:
    planilha = workbook[nome_planilha]
except KeyError:
    print(f"A planilha '{nome_planilha}' não foi encontrada no arquivo do Excel.")
    workbook.close()
    exit()

# Especificar as colunas a serem mantidas
colunas_manter = [1, 2, 3, 5, 11, 13]  # Substitua pelos índices das colunas que deseja manter

# Chamar a função para manter apenas as colunas desejadas
manter_colunas(planilha, colunas_manter)

# Salvar as modificações no arquivo
arquivo_modificado = r'C:\Users\venda\OneDrive\Área de Trabalho\Resolver Planilha\estoque_simplificada.xlsx'
workbook.save(arquivo_modificado)

# Fechar o arquivo do Excel
workbook.close()