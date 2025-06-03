import pandas as pd
import numpy as np

file_path = r'C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\analise_topdown_pastagem_jun25_correta.xlsx'

sheet_names = [
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
    'IMAZETAPIR 100 G L',
    'LAMBDA-CIALOTRINA + TIAMETOXAM',
    'LAMBDA-CIALOTRINA 250 G L',
    'MESOTRIONA 480 G L',
    'METOMIL 216 G L',
    'METSULFUROM 600 G Kg'
]

filiais = [
    'AGUA BOA', 'ARAGUAINA', 'BARRA DO GARÇAS', 'CACERES', 'Campo Grande HUB',
    'COLINAS DO TOCANTINS', 'CONFRESA', 'GOIANIA', 'GURUPI', 'Jaru',
    'Ji Paraná', 'JUARA', 'JUSSARA', 'MARABA', 'NOVA BANDEIRANTES',
    'Ouro Preto do Oeste', 'PARAISO DO TOCANTINS', 'Pimenta Bueno',
    'PONTES E LACERDA', 'Porto Velho', 'Rio Branco', 'SAO MIGUEL DO ARAGUAIA',
    'VILA RICA', 'Vilhena'
]

colunas_real_vendido_letras = [
    'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
    'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK'
]

colunas_forecast_letras = [
    'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA',
    'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM'
]

def letras_para_indices(letras):
    resultado = []
    for letra in letras:
        total = 0
        for char in letra:
            total = total * 26 + (ord(char.upper()) - ord('A') + 1)
        resultado.append(total - 1)
    return resultado

indices_real_vendido = letras_para_indices(colunas_real_vendido_letras)
indices_forecast = letras_para_indices(colunas_forecast_letras)

indice_campanha = letras_para_indices(['M'])[0]
indice_data = letras_para_indices(['J'])[0]

dfs = []

for sheet in sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet, header=9, nrows=42)
    
    campanha_col = df.iloc[:, indice_campanha].astype(str)
    data_col = df.iloc[:, indice_data].astype(str)
    
    real_vendido = df.iloc[:, indices_real_vendido]
    forecast = df.iloc[:, indices_forecast]
    
    real_vendido.columns = filiais
    forecast.columns = filiais

    # Formato longo
    real_vendido_long = real_vendido.melt(var_name='Filial Planejamento', value_name='Real Vendido').reset_index(drop=True)
    forecast_long = forecast.melt(var_name='Filial Planejamento', value_name='Previsão').reset_index(drop=True)
    forecast_long['Previsão'] = pd.to_numeric(forecast_long['Previsão'], errors='coerce').fillna(0).round().astype(int)
    
    # Use np.tile para repetir cada valor de campanha/data para todas as filiais daquela linha
    campanha_tiled = np.tile(campanha_col.values, len(filiais))
    data_tiled = np.tile(data_col.values, len(filiais))
    
    df_combined = pd.DataFrame({
        'Princípio Ativo': sheet,
        'Campanha': campanha_tiled,
        'Filial Planejamento': real_vendido_long['Filial Planejamento'],
        'Real Vendido': real_vendido_long['Real Vendido'],
        'Previsão': forecast_long['Previsão'],
        'Data': data_tiled
    })

    dfs.append(df_combined)

df_final = pd.concat(dfs, ignore_index=True)

df_final = df_final[['Princípio Ativo', 'Campanha', 'Filial Planejamento', 'Real Vendido', 'Previsão', 'Data']]

df_final.to_excel(r'C:\Users\BrunoAriasRocha\Downloads\resultado_final_com_campanha.xlsx', index=False)

print("Arquivo 'resultado_final_com_campanha.xlsx' salvo com sucesso!")