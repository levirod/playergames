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
    
    # Salvar as modificações no arquivo
    planilha.to_excel(nome_arquivo, index=False)
    
    print("Modificações concluídas.")

# Exemplo de uso
arquivo = r'C:\Users\venda\OneDrive\Área de Trabalho\Resolver Planilha\teste_preenchido.xlsx'  # Substitua pelo caminho correto do seu arquivo
modificar_planilha(arquivo)
