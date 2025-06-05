import pandas as pd
import numpy as np

file_path = r'C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\analise_topdown__adubos_pastagem_jun25.xlsx'

# Lê dinamicamente os nomes das abas do arquivo Excel
xls = pd.ExcelFile(file_path)
sheet_names = xls.sheet_names  # lista com todos os nomes das abas

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
    try:
        # Lê a aba pelo nome, com header na linha 10 (índice 9), 42 linhas de dados
        df = pd.read_excel(xls, sheet_name=sheet, header=9, nrows=42)

        # Checa se há colunas suficientes para Campanha e Data
        if df.shape[1] <= max(indice_campanha, indice_data):
            print(f"Aba '{sheet}' não possui as colunas necessárias (Campanha ou Data). Pulando.")
            continue

        # Checa se há colunas suficientes para Real Vendido e Forecast
        if df.shape[1] <= max(max(indices_real_vendido), max(indices_forecast)):
            print(f"Aba '{sheet}' não possui todas as colunas de Real Vendido ou Forecast. Pulando.")
            continue

        campanha_col = df.iloc[:, indice_campanha].astype(str)
        data_col = df.iloc[:, indice_data].astype(str)

        real_vendido = df.iloc[:, indices_real_vendido]
        forecast = df.iloc[:, indices_forecast]

        real_vendido.columns = filiais
        forecast.columns = filiais

        real_vendido_long = real_vendido.melt(var_name='Filial Planejamento', value_name='Real Vendido').reset_index(drop=True)
        forecast_long = forecast.melt(var_name='Filial Planejamento', value_name='Previsão').reset_index(drop=True)
        forecast_long['Previsão'] = pd.to_numeric(forecast_long['Previsão'], errors='coerce').fillna(0).round().astype(int)

        # Repete campanha e data para cada filial da linha
        campanha_tiled = np.tile(campanha_col.values, len(filiais))
        data_tiled = np.tile(data_col.values, len(filiais))

        df_combined = pd.DataFrame({
            'Princípio Ativo': sheet,  # Usa o nome da aba como princípio ativo
            'Campanha': campanha_tiled,
            'Filial Planejamento': real_vendido_long['Filial Planejamento'],
            'Real Vendido': real_vendido_long['Real Vendido'],
            'Previsão': forecast_long['Previsão'],
            'Data': data_tiled
        })

        dfs.append(df_combined)
    except Exception as e:
        print(f"Erro ao processar a aba '{sheet}': {e}")
        continue

df_final = pd.concat(dfs, ignore_index=True)

df_final = df_final[['Princípio Ativo', 'Campanha', 'Filial Planejamento', 'Real Vendido', 'Previsão', 'Data']]

df_final.to_excel(r'C:\Users\BrunoAriasRocha\Downloads\resultado_final_com_campanha_adubos_pastagem.xlsx', index=False)

print("Arquivo salvo com sucesso!")
