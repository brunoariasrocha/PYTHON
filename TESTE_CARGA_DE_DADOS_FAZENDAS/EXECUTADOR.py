import pandas as pd


#Lendo arquivo excel
caminho_arquivo = r'C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Área de Trabalho\Teste Carga de Dados Fazenda\CARGA_FAZENDA.xlsx'
dados_excel = pd.read_excel(caminho_arquivo)

#Selecionar Coluna
colunas_selecionadas = ['CODCLIENTE', 'NOMEFAZENDA']
print(dados_excel[colunas_selecionadas])


#Imprime todos os dados da planilha
#print(dados_excel)