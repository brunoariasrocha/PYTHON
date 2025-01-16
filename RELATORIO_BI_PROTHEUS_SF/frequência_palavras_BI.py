import pandas as pd
import re
from collections import Counter

# Carregue seus dados de chamados em um DataFrame usando pandas.
# df = pd.read_excel(r"C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\RELATÓRIOS SEMANAIS DE TICKETS\RELATORIOS_BI\BASE_RELATÓRIOS_SEMANAIS_PROTHEUS\BASE1.xlsx")

df = dataset
descricoes = df['Descrição']

# Limpa o texto removendo pontuações, números e convertendo para minúsculas.
def limpar_texto(texto):
    texto = texto.lower()
    texto = re.sub(r'[^\w\s]', '', texto)
    texto = re.sub(r'\d+', '', texto)
    return texto

df['descricao_limpa'] = df['Descrição'].apply(limpar_texto)

# Tokenização das palavras (simplesmente separando por espaços)
df['tokens'] = df['descricao_limpa'].apply(lambda x: x.split())

# Lista manual de stopwords (português)
stop_words = {
    'a', 'à', 'às', 'agora', 'aí', 'ainda', 'além', 'algum', 'alguma', 'algumas', 'alguns', 'ante', 'antes', 'ao', 
    'aos', 'após', 'aquela', 'aquelas', 'aquele', 'aqueles', 'aquilo', 'as', 'até', 'baixo', 'bem', 'bom', 'cada', 
    'cá', 'coisa', 'com', 'como', 'contra', 'contudo', 'da', 'daquela', 'daquelas', 'daquele', 'daqueles', 'dar', 
    'das', 'de', 'dela', 'delas', 'dele', 'deles', 'demais', 'dentro', 'depois', 'desde', 'dessa', 'dessas', 
    'desse', 'desses', 'desta', 'destas', 'deste', 'destes', 'disso', 'disto', 'do', 'dos', 'e', 'é', 'ela', 
    'elas', 'ele', 'eles', 'em', 'enquanto', 'entre', 'era', 'essa', 'essas', 'esse', 'esses', 'esta', 'estas', 
    'este', 'estes', 'eu', 'foi', 'for', 'fora', 'hoje', 'já', 'lá', 'lhe', 'lhes', 'mais', 'mas', 'me', 'mesmo', 
    'meu', 'minha', 'minhas', 'meus', 'na', 'nas', 'nem', 'nenhum', 'nessa', 'nessas', 'nesse', 'nesses', 'nesta', 
    'nestas', 'neste', 'nestes', 'ninguém', 'no', 'nos', 'nós', 'nossa', 'nossas', 'nosso', 'nossos', 'num', 
    'numa', 'nunca', 'o', 'os', 'ou', 'para', 'por', 'porque', 'qual', 'qualquer', 'quando', 'quanto', 'que', 
    'quem', 'se', 'sem', 'ser', 'seu', 'seus', 'só', 'sob', 'sobre', 'sua', 'suas', 'tal', 'também', 'tampouco', 
    'te', 'tem', 'tendo', 'ter', 'teu', 'teus', 'toda', 'todas', 'todo', 'todos', 'tu', 'tua', 'tuas', 'tudo', 
    'um', 'uma', 'umas', 'uns', 'você', 'vocês'
}

# Remover stopwords
df['tokens'] = df['tokens'].apply(lambda x: [word for word in x if word not in stop_words])

# Conte a frequência de cada palavra.
todas_palavras = [palavra for tokens in df['tokens'] for palavra in tokens]
contador = Counter(todas_palavras)
palavras_mais_comuns = contador.most_common(10)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Gravando no Excel:
df_palavras = pd.DataFrame(palavras_mais_comuns, columns=['Palavra', 'Frequência'])

# Salve o DataFrame em um arquivo Excel.
# caminho_arquivo = r'C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\RELATÓRIOS SEMANAIS DE TICKETS\RELATORIOS_BI\BASE_RELATÓRIOS_SEMANAIS_PROTHEUS\resultado_palavras_frequentes.xlsx'

# with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
#     df_palavras.to_excel(writer, sheet_name='Palavras Frequentes', index=False)
