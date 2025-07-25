import pandas as pd
from prophet import Prophet
from datetime import datetime, timedelta

def proximo_dia_mes(data_ref, dia_alvo):
    """
    Calcula a próxima data do calendário que seja o dia 'dia_alvo' 
    no mês corrente ou no próximo mês se já passou.
    Trata meses com diferentes números de dias.
    """
    ano = data_ref.year
    mes = data_ref.month
    dia = dia_alvo

    # Se dia alvo já passou neste mês, avança um mês
    if data_ref.day >= dia_alvo:
        if mes == 12:
            mes = 1
            ano += 1
        else:
            mes += 1

    try:
        return datetime(ano, mes, dia)
    except ValueError:
        # Caso dia_alvo inválido para o mês (ex: 31 em fevereiro)
        # Retorna o último dia do mês
        proximo_mes = mes % 12 + 1
        ano_proximo_mes = ano + 1 if proximo_mes == 1 else ano
        next_month_first = datetime(ano_proximo_mes, proximo_mes, 1)
        ultimo_dia = next_month_first - timedelta(days=1)
        return ultimo_dia

# Caminho para o arquivo base2.csv - ajuste conforme seu ambiente
caminho_arquivo = r"C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\VSCode_Arquivos\ML_DATAFORNECEDOR\base2.csv"

# Lê a base com tratamento da data no formato DD/MM/YYYY brasileiro
df = pd.read_csv(
    caminho_arquivo,
    sep=';',
    parse_dates=['data_emissao'],
    date_parser=lambda x: pd.to_datetime(x, format='%d/%m/%Y')
)

resultados = []
hoje = pd.Timestamp.today().normalize()

for fornecedor in df['fornecedor'].unique():
    dados = df[df['fornecedor'] == fornecedor].copy()
    dados.rename(columns={'data_emissao': 'ds'}, inplace=True)
    
    # Agrupar por data, contando o número de eventos (notas)
    dados_agg = dados.groupby('ds').size().reset_index(name='y')
    dados_agg = dados_agg.sort_values('ds')

    if len(dados_agg) == 0:
        resultados.append({
            'fornecedor': fornecedor,
            'proxima_nf': None,
            'dias_faltando': None,
            'observacao': 'Sem dados'
        })
        continue

    # Extrair o dia do mês das datas históricas
    dias_mes = dados_agg['ds'].dt.day

    # Calcular o dia mais frequente da emissão (moda)
    dia_frequente = dias_mes.mode()
    if len(dia_frequente) > 0:
        dia_alvo = int(dia_frequente.iloc[0])
    else:
        # Se não conseguir detectar a moda, usa o dia da última emissão
        dia_alvo = dados_agg['ds'].iloc[-1].day

    # Calcular a próxima data provável baseada na periodicidade mensal
    proxima_data_calendario = proximo_dia_mes(hoje.to_pydatetime(), dia_alvo)

    # Estimar a previsão com Prophet para comparação
    try:
        if len(dados_agg) >= 2:
            model = Prophet()
            model.fit(dados_agg[['ds', 'y']])
            future = model.make_future_dataframe(periods=60)
            forecast = model.predict(future)
            previsao_futura = forecast[forecast['ds'] > dados_agg['ds'].max()]
            proxima_data_prophet = previsao_futura[previsao_futura['yhat'] > 0.9]['ds'].min()
            if pd.isnull(proxima_data_prophet):
                proxima_data_prophet = previsao_futura['ds'].min()
        else:
            proxima_data_prophet = None
    except Exception as e:
        proxima_data_prophet = None

    # Escolher a data futura mais próxima entre as duas previsões válidas
    datas_possiveis = [d for d in [proxima_data_calendario, proxima_data_prophet] if d is not None and d >= hoje]
    if datas_possiveis:
        proxima_final = min(datas_possiveis)
    else:
        proxima_final = proxima_data_calendario  # fallback

    dias_faltando = (proxima_final - hoje).days

    resultados.append({
        'fornecedor': fornecedor,
        'proxima_nf': proxima_final.date(),
        'dias_faltando': dias_faltando,
        'observacao': 'Previsão baseada em dia frequente do mês + Prophet'
    })

# Criar DataFrame para salvar
df_resultados = pd.DataFrame(resultados)

# Formatar data no padrão DD-MM-YYYY
df_resultados['proxima_nf'] = pd.to_datetime(df_resultados['proxima_nf']).dt.strftime('%d-%m-%Y')

# Caminho para salvar o Excel com o resultado - ajuste conforme seu ambiente
caminho_saida = r"C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\VSCode_Arquivos\ML_DATAFORNECEDOR\previsao_nfs.xlsx"

# Salvar no Excel
df_resultados.to_excel(caminho_saida, index=False)

print("✅ Previsões calibradas baseadas na periodicidade mensal e Prophet foram salvas em Excel.")
