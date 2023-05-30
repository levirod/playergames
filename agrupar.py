import pandas as pd
import locale

# Definir a formatação local para lidar com o separador decimal e ponto de milhar
locale.setlocale(locale.LC_ALL, '')

# Carregar a planilha
planilha = pd.read_excel(r'C:\Users\venda\OneDrive\Área de Trabalho\Resolver Planilha\planilha_modificada.xlsx')

# Converter a coluna "Quantidade em estoque" para numérico
planilha['Quantidade em estoque'] = planilha['Quantidade em estoque'].apply(
    lambda x: locale.atof(str(x)) if pd.notnull(x) else x
)

# Remover linhas duplicadas e atualizar a coluna "Quantidade em estoque"
planilha_agrupada = planilha.groupby('Descrição do Produto').agg({
    'Quantidade em estoque': 'sum',
    'Valor Venda': 'first'  # Manter o primeiro valor de venda encontrado
}).reset_index()

# Salvar a planilha atualizada
planilha_agrupada.to_excel(r'C:\Users\venda\OneDrive\Área de Trabalho\Resolver Planilha\planilha_atualizada.xlsx', index=False)
