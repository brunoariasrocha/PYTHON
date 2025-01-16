import pandas as pd
import re

#Importando tabela e gravando na variável tabela
tabela = pd.read_excel(r"C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\FILTRADOS_SEM_ABC.xlsx")

#Gravando a tabela em um DataFrame
df = pd.DataFrame(tabela)

#Salvando a coluna E-mail em uma variável
coluna_codigocli = df[['Codigo']]
coluna_email = df[['E-Mail']]
coluna_loja = df[['Loja']]
coluna_nome = df[['Nome']]

#Traz os caracteres que quero achar
regex_pattern = r'@[^a-zA-Z]|(\.com|\.com\.br)[^ ]|;$|\.com\w{0,2}\b|\.br\w{0,1}\b|^$|^[^a-zA-Z0-9]+$|^[0-9]+$'

# Cria uma lista para armazenar os e-mails correspondentes
emails_correspondentes = []
codigos_correspondentes = []
lojas_correspondentes = []
nomes_correspondentes = []

for index, row in df.iterrows():
    email = row['E-Mail']
    codigo = row['Codigo']
    lojas = row['Loja']
    nomes = row['Nome']

    if pd.notna(email) and re.search(regex_pattern, email):
        emails_correspondentes.append(email)
        codigos_correspondentes.append(codigo)
        lojas_correspondentes.append(lojas)
        nomes_correspondentes.append(nomes)

# Criar um novo DataFrame a partir da lista de e-mails correspondentes
df_emails_correspondentes = pd.DataFrame({'E-Mail': emails_correspondentes, 'Codigo' : codigos_correspondentes, 'Loja' : lojas_correspondentes, 'Nome' : nomes_correspondentes})

# Salvar o DataFrame em um arquivo Excel com o caminho completo
df_emails_correspondentes.to_excel(r"C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\emails_correspondentes2.xlsx", engine='openpyxl', index=False)

#print(coluna_email)
#print(df[['E-Mail']])
#print(tabela)