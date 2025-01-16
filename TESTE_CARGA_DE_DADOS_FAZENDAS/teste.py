import pandas as pd

# Carregue o arquivo da planilha
caminho_arquivo = r'C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Área de Trabalho\Teste Carga de Dados Fazenda\CARGA_FAZENDA.xlsx'
df = pd.read_excel(caminho_arquivo)

# Realize a correspondência entre as colunas 'Coluna1' e 'Coluna2' e encontre um valor único
valor_procurado = 'valor_procurado'
valor_unico = df.loc[caminho_arquivo['NOMEFAZENDA'] == valor_procurado, 'CODCLIENTE'].iloc[0]


print(valor_unico)