import pandas as pd
import re
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from collections import Counter
from wordcloud import WordCloud
import matplotlib.pyplot as plt



#Carregue seus dados de chamados em um DataFrame usando pandas.
df = pd.read_excel(r"C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\RELATÓRIOS SEMANAIS DE TICKETS\RELATORIOS_BI\BASE_RELATÓRIOS_SEMANAIS_PROTHEUS\BASE1.xlsx")

descricoes = df['Descrição']

#Limpa o texto removendo pontuações, números, e convertendo para minúsculas.
def limpar_texto(texto):
    texto = texto.lower()
    texto = re.sub(r'[^\w\s]', '', texto)
    texto = re.sub(r'\d+', '', texto)
    return texto

df['descricao_limpa'] = df['Descrição'].apply(limpar_texto)

#Tokenização das palavras
df['tokens'] = df['descricao_limpa'].apply(word_tokenize)

#Remova palavras comuns e irrelevantes como "e", "de", "o", etc.
stop_words = set(stopwords.words('portuguese'))
df['tokens'] = df['tokens'].apply(lambda x: [word for word in x if word not in stop_words])

#Conte a frequência de cada palavra.
todas_palavras = [palavra for tokens in df['tokens'] for palavra in tokens]
contador = Counter(todas_palavras)
palavras_mais_comuns = contador.most_common(10)

#Nuvem de Palavras
#texto = ' '.join(todas_palavras)
#nuvem_palavras = WordCloud(width=800, height=400, max_font_size=100, background_color='white').generate(texto)

#plt.figure(figsize=(10, 5))
#plt.imshow(nuvem_palavras, interpolation='bilinear')
#plt.axis('off')
#plt.show()

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------
#Gráfico de Barras
#palavras, frequencias = zip(*palavras_mais_comuns)
#plt.bar(palavras, frequencias)
#plt.show()

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Gravando no Excel:
 #Supondo que 'contador' seja o Counter com as frequências das palavras
palavras_mais_comuns = contador.most_common(10)  #<- Por exemplo, as 10 palavras mais comuns
df_palavras = pd.DataFrame(palavras_mais_comuns, columns=['Palavra', 'Frequência'])

#Salve o DataFrame em um arquivo Excel.
caminho_arquivo = r'C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\RELATÓRIOS SEMANAIS DE TICKETS\RELATORIOS_BI\BASE_RELATÓRIOS_SEMANAIS_PROTHEUS\resultado_palavras_frequentes.xlsx'

with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
    df_palavras.to_excel(writer, sheet_name='Palavras Frequentes', index=False)