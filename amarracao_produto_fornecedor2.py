import pandas as pd

# Carregar o arquivo Excel com os dados originais
input_file = r"C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\PROJETO LAVORO\FORNECEDORES\combinacoes_produtos_fornecedores2.xlsx"
data = pd.read_excel(input_file, sheet_name='Planilha1')

# Gerar uma lista com as 30 lojas (Loja 00 até Loja 29)
stores = [f"Loja {str(i).zfill(2)}" for i in range(30)]

# Replicar cada produto para todas as lojas
expanded_data = pd.concat(
    [data.assign(A5_LOJA=store) for store in stores],
    ignore_index=True
)

# Salvar o resultado em um novo arquivo Excel
output_file = r"C:\Users\BrunoAriasRocha\OneDrive - AgroOpps\Documentos\PROJETO LAVORO\FORNECEDORES\testecarga.xlsx"
expanded_data.to_excel(output_file, index=False)

print(f"Arquivo salvo em: {output_file}")
