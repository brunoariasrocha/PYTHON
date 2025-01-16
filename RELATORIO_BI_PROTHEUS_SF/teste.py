import pandas as pd


#Importando tabela e gravando na variável tabela
layout01 = pd.read_excel(r"C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\VSCode_Arquivos\RELATORIO_BI_PROTHEUS_SF\RELATORIOS\Relatório_Service_Desk.xls")

# Garantindo que a coluna 'Assign To Individual' seja do tipo string e removendo espaços em branco
layout01['Assign To Individual'] = layout01['Assign To Individual'].astype(str).str.strip()

# Lista de nomes para filtrar
nomes = ['Bruno', 'Robson', 'Alex', 'Juliana']

# Criando a expressão regular para buscar qualquer um dos nomes
regex_pattern = '|'.join(nomes)

# Aplicando o filtro para encontrar linhas que contenham 'Bruno' ou 'Robson' na coluna 'Assign To Individual'
layout01_filtrado = layout01[layout01['Assign To Individual'].str.contains(regex_pattern, na=False, case=False)]

#Arquivo Novo
layout01_filtrado.to_excel(r'C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\VSCode_Arquivos\RELATORIO_BI_PROTHEUS_SF\BASE_BI\base_bi.xlsx', index=False)