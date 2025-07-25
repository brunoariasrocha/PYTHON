import pandas as pd
import openpyxl as px

arquivopadrao = pd.read_excel("C:\\Users\\BrunoAriasRocha\\OneDrive - AgroOpps\\Documentos\\PROJETO LAVORO\\CONTAS_A_RECEBER\\Base Projeto 500 Axia (13Mai25) vs1.xlsb", sheet_name='Planilha1')

print(arquivopadrao)