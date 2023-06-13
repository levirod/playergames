import openpyxl

def alterar_primeira_letra_maiusculo(coluna):
    # Abre o arquivo do Excel
    caminho_arquivo = 'planilha_tagplus.xlsx'  # Substitua pelo caminho do seu arquivo Excel
    livro = openpyxl.load_workbook(caminho_arquivo)

    # Seleciona a planilha desejada
    nome_planilha = 'Sheet1'  # Substitua pelo nome da sua planilha
    planilha = livro[nome_planilha]

    # Itera pelas células na coluna específica e altera a primeira letra de cada palavra para maiúsculo
    for celula in planilha[coluna]:
        valor = celula.value
        if valor is not None:  # Verifica se a célula não está vazia
            palavras = valor.split()  # Separa as palavras na célula
            palavras_maiusculo = [palavra.capitalize() for palavra in palavras]  # Altera a primeira letra de cada palavra para maiúsculo
            novo_valor = ' '.join(palavras_maiusculo)  # Junta as palavras modificadas
            celula.value = novo_valor

    # Salva as alterações no arquivo
    livro.save(caminho_arquivo)

    # Fecha o arquivo
    livro.close()

# Chama a função para alterar a coluna 'A'
alterar_primeira_letra_maiusculo('A')
