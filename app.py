with abas[0]:
    st.title("Saldo Produção")

    file = st.file_uploader("PDF Saldo", type=["pdf"])

    if file:
        with open("saldo_temp.pdf", "wb") as f:
            f.write(file.read())

        if st.button("Processar Saldo"):
            linhas = []

            with pdfplumber.open("saldo_temp.pdf") as pdf:
                for p in pdf.pages:
                    texto = p.extract_text()
                    if texto:
                        linhas.extend(texto.split("\n"))

            dados = {}
            codigo_atual = None

            for linha in linhas:
                linha = linha.strip()
                if not linha:
                    continue

                if "ALMOXARIFADO" in linha.upper():
                    codigo_match = None
                else:
                    codigo_match = re.search(r"\b([A-Z]{1,3}\d{3,5})\b", linha)

                if codigo_match:
                    codigo_atual = codigo_match.group(1)
                    if codigo_atual not in dados:
                        dados[codigo_atual] = {
                            "Codigo": codigo_atual,
                            "Saldo Total": 0,
                            "Saldo Almox 3": 0,
                            "Saldo Almox 30": 0
                        }
                    continue

                if "ALMOXARIFADO" in linha.upper() and codigo_atual:
                    nums = re.findall(r"[\d\.]+\,\d+", linha)
                    if nums:
                        valor = float(nums[-1].replace(".", "").replace(",", "."))
                        dados[codigo_atual]["Saldo Total"] += valor

                        if re.search(r"ALMOXARIFADO\s*:\s*30\b", linha.upper()):
                            dados[codigo_atual]["Saldo Almox 30"] += valor

                        if re.search(r"ALMOXARIFADO\s*:\s*3\b", linha.upper()):
                            dados[codigo_atual]["Saldo Almox 3"] += valor

            df = pd.DataFrame(dados.values())
            df["Data Processamento"] = agora().strftime("%d/%m/%Y")
            df["Hora Processamento"] = agora().strftime("%H:%M:%S")

            df.to_csv("saldo.csv", index=False)

            st.success("Saldo processado com sucesso.")
            st.dataframe(df, use_container_width=True)
