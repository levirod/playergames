import pandas as pd

def modificar_planilha(nome_arquivo):
    # Carregar a planilha existente
    planilha = pd.read_excel(nome_arquivo)
    
    # Modificar os valores das colunas
    planilha['tipo'] = 'sem-variação'
    planilha['ativo'] = 'S'
    planilha['usado'] = 'N'
    planilha['destaque'] = 'N'
    planilha['estoque-gerenciado'] = 'S'
    planilha['estoque-situacao-em-estoque'] = 'imediata'
    planilha['estoque-situacao-sem-estoque'] = 'indisponivel'
    planilha['preco-sob-consulta'] = 'N'
    

    # Salvar as modificações no arquivo
    planilha.to_excel(nome_arquivo, index=False)
    
    print("Modificações concluídas.")

# Exemplo de uso
arquivo = 'paraosite.xlsx'  # Substitua pelo caminho correto do seu arquivo
modificar_planilha(arquivo)
