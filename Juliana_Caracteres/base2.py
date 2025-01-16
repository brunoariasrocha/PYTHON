from fuzzywuzzy import fuzz
import pandas as pd

# Função para calcular similaridade entre uma palavra e uma lista de palavras
def encontrar_melhor_correspondencia(palavra, lista_palavras):
    # Converter para string e filtrar valores NaN
    palavra = str(palavra) if pd.notnull(palavra) else ""
    similaridade_maxima = 0
    palavra_mais_parecida = ""
    for p in lista_palavras:
        p = str(p) if pd.notnull(p) else ""
        similaridade = fuzz.ratio(palavra, p)
        if similaridade > similaridade_maxima:
            similaridade_maxima = similaridade
            palavra_mais_parecida = p
    return palavra_mais_parecida, similaridade_maxima

# Carregar dados do Excel
df = pd.read_excel(r'C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\ANALISE_RAYANE_QBR\COMPARA_PROD.xlsx')

# Aplicar a função para encontrar a melhor correspondência para cada valor de 'Palavra1' na coluna 'Palavra2'
df['Melhor Correspondência'], df['Similaridade Máxima'] = zip(*df['Palavra1'].apply(lambda x: encontrar_melhor_correspondencia(x, df['Palavra2'])))

# Escrever resultados em um novo arquivo Excel
df.to_excel(r'C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\ANALISE_RAYANE_QBR\resultado_da_comparacao_1.xlsx', index=False)
