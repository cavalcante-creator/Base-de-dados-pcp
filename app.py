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
                "Saldo Almox 3": 0
            }
        continue

    if "ALMOXARIFADO" in linha.upper() and codigo_atual:
        nums = re.findall(r"[\d\.]+\,\d+", linha)
        if nums:
            valor = float(nums[-1].replace(".", "").replace(",", "."))
            dados[codigo_atual]["Saldo Total"] += valor
            if re.search(r"ALMOXARIFADO\s*:\s*3\b", linha.upper()):
                dados[codigo_atual]["Saldo Almox 3"] += valor
