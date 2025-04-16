import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import re
import time

# Produtos a buscar
produtos = [
    "ivomec gold 500ml",
    # Adicione mais produtos aqui
]

# Lista para armazenar os resultados
dados_resultado = []

# Função para extrair códigos de barras (GTIN/EAN)
def extrair_codigo_barras(conteudo_html):
    padroes = [
        r"GTIN(?:\/EAN)?:?\s*(\d{8,14})",
        r"EAN(?:-13)?:?\s*(\d{8,14})",
        r"Código\s+de\s+Barras\s*:?[\s\-]*(\d{8,14})"
    ]
    for padrao in padroes:
        resultado = re.search(padrao, conteudo_html, re.IGNORECASE)
        if resultado:
            return resultado.group(1)
    return None

# Caminho para o perfil do Chrome (opcional)
user_data_dir = r"C:\Users\BrunoAriasRocha\AppData\Local\Google\Chrome\User Data"
profile_dir = "Default"

# Configuração do navegador com perfil e stealth
options = uc.ChromeOptions()
options.user_data_dir = user_data_dir
options.add_argument(f"--profile-directory={profile_dir}")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920x1080")
# options.add_argument("--headless")  # Ative se quiser rodar sem abrir o navegador

# Inicializa o navegador
driver = uc.Chrome(options=options)

# Loop de busca
for produto in produtos:
    print(f"\n🔎 Buscando produtos semelhantes a: {produto}")

    driver.get('https://cosmos.bluesoft.com.br/buscar_produtos_por_codigo_de_barras')

    try:
        # Espera até o campo de busca estar presente
        campo_busca = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.ID, 'search-input'))
        )
        campo_busca.clear()
        campo_busca.send_keys(produto)
        campo_busca.send_keys(Keys.RETURN)

        # Aguarda o carregamento da próxima página
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )

        # Localiza o elemento com o GTIN/EAN diretamente
        try:
            elemento_gtin = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//strong[contains(text(),'GTIN/EAN:')]"))
            )
            codigo_gtin = elemento_gtin.text.replace("GTIN/EAN:", "").strip()
            print(f"✅ Código(s) encontrado(s): {codigo_gtin}")
        except:
            codigo_gtin = "Não encontrado"
            print("⚠️ Nenhum código encontrado.")

        dados_resultado.append({
            "Produto": produto,
            "Código(s) de Barras": codigo_gtin,
            "Fonte": 'https://cosmos.bluesoft.com.br/buscar_produtos_por_codigo_de_barras'
        })

    except Exception as e:
        print(f"❌ Erro ao buscar produto: {e}")
        # Salva HTML para debug
        with open(f"erro_html_{produto.replace(' ', '_')}.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)

    # Atraso para evitar bloqueios
    time.sleep(2)

# Fechar o navegador
driver.quit()

# Exportar resultados
df_resultados = pd.DataFrame(dados_resultado)
arquivo_excel = r"C:\Users\BrunoAriasRocha\Downloads\codigos_barras_produtos_chrome.xlsx"
df_resultados.to_excel(arquivo_excel, index=False)

print(f"\n📁 Resultado exportado com sucesso para: {arquivo_excel}")
