import pandas as pd
import locale

# Definir a formatação local para lidar com o separador decimal e ponto de milhar
locale.setlocale(locale.LC_ALL, '')

# Carregar a planilha
planilha = pd.read_excel('planilha_tagplus.xlsx')

# Converter a coluna "Quantidade em estoque" para numérico
planilha['Quantidade em estoque'] = planilha['Quantidade em estoque'].apply(
    lambda x: locale.atof(str(x)) if pd.notnull(x) else x
)

# Remover linhas duplicadas e atualizar a coluna "Quantidade em estoque"
planilha_agrupada = planilha.groupby('Descrição do Produto').agg({
    'Quantidade em estoque': 'sum',
    'Código interno': 'first', # Manter o primeiro valor de venda encontrado
    'Código de Barras': 'first',
    'Valor Venda': 'first',
    'Código NCM': 'first'
    
}).reset_index()

# Salvar a planilha atualizada
planilha_agrupada.to_excel('planilha_tagplus.xlsx', index=False)
