import pandas as pd
import re

#Importando tabela e gravando na variável tabela
tabela = pd.read_excel(r"C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\LOTES_BASE.xlsx")

#Gravando a tabela em um DataFrame
df = pd.DataFrame(tabela)

#Salvando a coluna Lotes em uma variável
coluna_lotes = df[['Lote']]

#Traz os caracteres que quero achar
regex_pattern = r'^\s*'

# Cria uma lista para armazenar os lotes correspondentes
lotes_correspondentes = []

for index, row in df.iterrows():
    lotes = row['Lote']

    if pd.notna(lotes) and re.search(regex_pattern, lotes):
        lotes_correspondentes.append(lotes)


# Criar um novo DataFrame a partir da lista de e-mails correspondentes
df_lotes_correspondentes = pd.DataFrame({'Lote': lotes_correspondentes})

# Salvar o DataFrame em um arquivo Excel com o caminho completo
df_lotes_correspondentes.to_excel(r"C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\lotes_correspondentes.xlsx", engine='openpyxl', index=False)