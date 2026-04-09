import pdfplumber
import pandas as pd
import re
import glob
import os

# ============================================================
# LOCALIZAR PDF MAIS RECENTE
# ============================================================

pasta_pdf = "."
padrao_pdf = os.path.join(pasta_pdf, "SALDO PRODUÇÃO *.pdf")

arquivos = glob.glob(padrao_pdf)

if not arquivos:
    print("❌ Nenhum PDF encontrado.")
    exit()

arquivos.sort(key=os.path.getmtime)
arquivo_pdf = arquivos[-1]

print(f"✅ Usando PDF: {arquivo_pdf}")

# ============================================================
# LER PDF
# ============================================================

linhas = []

with pdfplumber.open(arquivo_pdf) as pdf:
    for pagina in pdf.pages:
        texto = pagina.extract_text()
        if texto:
            linhas.extend(texto.split("\n"))

# ============================================================
# EXTRAÇÃO
# ============================================================

dados = {}
codigo_atual = None

for linha in linhas:

    linha = linha.strip()
    linha_upper = linha.upper()

    # detectar código do item
    codigo_match = re.search(r'\b[A-Z]{2}\d{4}\b', linha)

    if codigo_match:
        codigo_atual = codigo_match.group()

        if codigo_atual not in dados:
            dados[codigo_atual] = {
                "Saldo Total": 0,
                "Saldo Almox 3": 0
            }

        continue

    # linhas de almoxarifado
    if "ALMOXARIFADO" in linha_upper and codigo_atual:

        numeros = re.findall(r'[\d\.]+\,\d+', linha)

        if numeros:

            saldo_str = numeros[-1]

            try:
                saldo = float(
                    saldo_str.replace('.', '').replace(',', '.')
                )
            except:
                saldo = 0

            # saldo total
            dados[codigo_atual]["Saldo Total"] += saldo

            # saldo almoxarifado 3
            if "ALMOXARIFADO: 3" in linha_upper:
                dados[codigo_atual]["Saldo Almox 3"] += saldo

# ============================================================
# DATAFRAME
# ============================================================

df = pd.DataFrame(
    [{"Codigo": k, **v} for k, v in dados.items()]
)

print("\n📊 Dados extraídos:")
print(df)

# ============================================================
# EXPORTAR
# ============================================================

saida = "saldo_producao_extraido.xlsx"
df.to_excel(saida, index=False)

print(f"\n✅ Arquivo gerado: {saida}")
