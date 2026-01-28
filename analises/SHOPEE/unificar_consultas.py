import os
import re
import pandas as pd


PASTA_DADOS = r"C:\Users\casad\Desktop\Caio\Python\analises\SHOPEE\saida\Caixa_2025"
ANO_PADRAO = 2025


ABA_FATO = "Lucro_por_SKU_PRECO_REAL"



SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_SAIDA = os.path.join(SCRIPT_DIR, "FATO_RESULTADO_SKU_MENSAL.xlsx")



MAPA_MESES = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4,
    "MAI": 5, "JUN": 6, "JUL": 7, "AGO": 8,
    "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12
}



def extrair_periodo(nome_arquivo):
    nome = nome_arquivo.upper()
    match = re.search(r"(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)", nome)

    if not match:
        raise ValueError(f"❌ Mês não identificado no nome do arquivo: {nome_arquivo}")

    mes_abrev = match.group(1)
    mes = MAPA_MESES[mes_abrev]
    ano = ANO_PADRAO
    periodo = f"{ano}-{str(mes).zfill(2)}"

    return ano, mes, periodo


def validar_colunas(df, colunas_obrigatorias, contexto):
    faltando = [c for c in colunas_obrigatorias if c not in df.columns]
    if faltando:
        raise ValueError(f"❌ Colunas faltando em {contexto}: {faltando}")



dfs = []

for arquivo in os.listdir(PASTA_DADOS):
    if not arquivo.endswith(".xlsx") or arquivo.startswith("~$"):
        continue

    caminho = os.path.join(PASTA_DADOS, arquivo)
    print(f"📂 Processando: {arquivo}")

    ano, mes, periodo = extrair_periodo(arquivo)

    df = pd.read_excel(caminho, sheet_name=ABA_FATO)
    df.columns = [c.strip() for c in df.columns]

    # contrato REAL da aba
    colunas_obrigatorias = [
        "SKU_x",
        "Produto",
        "unidades_pagas",
        "custo_total_unit",
        "cmv_total_sku",
        "preco_shopee",
        "receita_real_sku",
        "taxas_variaveis_sku",
        "taxa_fixa_sku",
        "lucro_sku",
        "margem_sku"
    ]

    validar_colunas(df, colunas_obrigatorias, f"{arquivo} → {ABA_FATO}")

    # normalização para padrão BI
    df = df.rename(columns={
        "SKU_x": "sku",
        "Produto": "produto",
        "unidades_pagas": "quantidade_vendida",
        "preco_shopee": "preco_unitario",
        "receita_real_sku": "receita_total",
        "custo_total_unit": "cmv_unitario",
        "cmv_total_sku": "cmv_total",
        "taxas_variaveis_sku": "taxas_variaveis",
        "taxa_fixa_sku": "taxas_fixas",
        "lucro_sku": "lucro_total",
        "margem_sku": "margem"
    })

    # metadados de período
    df["ano"] = ano
    df["mes"] = mes
    df["periodo"] = periodo
    df["origem_arquivo"] = arquivo

    # tipos 
    df["sku"] = df["sku"].astype(str)
    df["quantidade_vendida"] = df["quantidade_vendida"].astype(int)
    df["preco_unitario"] = df["preco_unitario"].astype(float)
    df["receita_total"] = df["receita_total"].astype(float)
    df["cmv_unitario"] = df["cmv_unitario"].astype(float)
    df["cmv_total"] = df["cmv_total"].astype(float)
    df["taxas_variaveis"] = df["taxas_variaveis"].astype(float)
    df["taxas_fixas"] = df["taxas_fixas"].astype(float)
    df["lucro_total"] = df["lucro_total"].astype(float)
    df["margem"] = df["margem"].astype(float)

    dfs.append(df)



fato_final = pd.concat(dfs, ignore_index=True)

# ordenação correta
fato_final = fato_final.sort_values(["ano", "mes", "sku"])

# validação de duplicidade (SKU + período)
duplicados = fato_final.duplicated(subset=["sku", "periodo"])
if duplicados.any():
    raise ValueError("❌ Existem SKUs duplicados no mesmo período. Corrija antes de seguir.")



fato_final.to_excel(ARQUIVO_SAIDA, index=False)

print("\n✅ PROCESSO FINALIZADO COM SUCESSO")
print(f"📍 Arquivo salvo em:\n{ARQUIVO_SAIDA}")
