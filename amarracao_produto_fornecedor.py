import pandas as pd

# Carregar os dados da planilha, mantendo zeros à esquerda
dados = pd.read_excel(
    r'C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\CARGA_AMARRACAO_TODOS_FORNECEDORES.xlsx',
    dtype={'B1_COD': str, 'A2_COD': str, 'A2_LOJA': str}
)

# Garantir que as colunas estão tratadas como texto
produtos = dados['B1_COD'].drop_duplicates().reset_index(drop=True)
fornecedores_lojas = dados[['A2_COD', 'A2_LOJA']].drop_duplicates()

# Gerar todas as combinações
resultado = produtos.to_frame(name='Produto').merge(fornecedores_lojas, how='cross')

# Salvar o resultado em um novo arquivo Excel, mantendo o formato de texto
resultado.to_excel(
    r'C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\combinacoes_produtos_fornecedores.xlsx', 
    index=False, 
    engine='openpyxl'
)

print("Combinações geradas com sucesso! O arquivo foi salvo como 'combinacoes_produtos_fornecedores.xlsx'.")
