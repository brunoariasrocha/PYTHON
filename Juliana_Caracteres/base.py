import pandas as pd

def caracteres_iguais_na_mesma_ordem(palavra1, palavra2):
    contagem = 0
    for letra1, letra2 in zip(palavra1, palavra2):
        if letra1 == letra2:
            contagem += 1
    return contagem

# Função para aplicar em todas as linhas
def contar_caracteres_iguais_em_todas_as_linhas(df):
    resultado = []
    for index, row in df.iterrows():
        quantidade_caracteres_iguais = caracteres_iguais_na_mesma_ordem(row['Palavra1'], row['Palavra2'])
        resultado.append(quantidade_caracteres_iguais)
    return resultado

# Carregar dados do Excel
df = pd.read_excel(r'C:\Users\BrunoAriasRocha\Downloads\TESTE_JULIANA.xlsx')

# Calcular resultados
df['Resultado'] = contar_caracteres_iguais_em_todas_as_linhas(df)

# Escrever resultados em um novo arquivo Excel
df.to_excel(r'C:\Users\BrunoAriasRocha\Downloads\resultados2.xlsx', index=False)
