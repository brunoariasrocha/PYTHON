import pandas as pd

#Lê Tabela
#df = pd.read_excel('C:\\Users\\BrunoAriasRocha\\OneDrive - AgroOpps\\Documentos\\analise_topdown_pastagem_jun25.xlsx')

#print(df.head())

import pandas as pd

# Lista com os princípios ativos
principios_ativos = [
    '2,4D 670 G L', 
    'ACETAMIPRIDO + BIFENTRINA 250 G',
    'ACETAMIPRIDO + CIPERMETRINA 100',
    'ATRAZINA + MESOTRIONA 500 G L +',
    'ATRAZINA 500 G L',
    'Cipermetrina 250 G L',
    'CIPROCONAZOL + TIAMETOXAM 300 G',
    'CLORPIRIFOS 480 G L',
    'DIQUATE 200 G L',
    'FIPRONIL 800 G Kg',
    'Glifosato 360 G L',
    'Glifosato 480 G L',
    'Glifosato 500 G L',
    'Glifosato 720 G Kg',
    'GLUFOSINATO 200 G L',
    'IMAZAPIQUE + IMAZAPIR 175 G Kg',
    'IMEZETAPIR 100 G L',
    'LAMBDA-CIALOTRINA + TIAMETOXAM',
    'LAMBDA-CIALOTRINA 250 G L',
    'MESOTRIONA 480 G L',
    'METOMIL 216 G L',
    'METSULFUROM 600 G Kg'
]

# Cria o DataFrame preenchendo as demais colunas com valores vazios
dados = {
    'Princípio Ativo': principios_ativos,
    'Filial Planejamento': [''] * len(principios_ativos),
    'Real Vendido': [None] * len(principios_ativos),
    'Forecast': [None] * len(principios_ativos),
    'Data': [''] * len(principios_ativos)
}

# Criar o DataFrame
df = pd.DataFrame(dados)

# Exibir resultado
print(df)
