import pandas as pd

# Ler a primeira planilha
planilha1 = pd.read_excel('planilha_tagplus.xlsx')

# Mapear os nomes das colunas
mapeamento_colunas = {
    'Descrição Do Produto': 'nome',
    'Quantidade em estoque': 'estoque-quantidade',
    'Código interno': 'sku',
    'Código de Barras': 'gtin',
    'Valor Venda': 'preco-cheio',
    'Código NCM': 'ncm'
}

# Renomear as colunas da primeira planilha
planilha1.rename(columns=mapeamento_colunas, inplace=True)

# Ler a segunda planilha
planilha2 = pd.read_excel('planilha-modelo.xlsx')

# Preencher a segunda planilha com os valores correspondentes da primeira planilha
planilha2['nome'] = planilha1['nome']
planilha2['estoque-quantidade'] = planilha1['estoque-quantidade']
planilha2['sku'] = planilha1['sku']
planilha2['gtin'] = planilha1['gtin']
planilha2['preco-cheio'] = planilha1['preco-cheio']
planilha2['ncm'] = planilha1['ncm']

# Salvar a segunda planilha com as colunas preenchidas
planilha2.to_excel('paraosite.xlsx', index=False)
