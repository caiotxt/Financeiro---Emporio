import os
import pandas as pd
import numpy as np



BASE_MESES = r"C:\Users\casad\Desktop\Caio\Python\analises\MERCADO LIVRE\meses"
DIMI_PATH = r"C:\Users\casad\Desktop\Caio\Python\analises\MERCADO LIVRE\costs\Planilha do Dimi 21-01.xlsx"
DIMI_SHEET = "PRODUTOS"
OUTDIR = r"C:\Users\casad\Desktop\Caio\Python\analises\MERCADO LIVRE\saida"
ANO = 2025

COL_PRODUTO = "Descrição da operação (reason)" 


MAPA_MESES = {
    "01": ("JAN", 1), "02": ("FEV", 2), "03": ("MAR", 3), "04": ("ABR", 4),
    "05": ("MAI", 5), "06": ("JUN", 6), "07": ("JUL", 7), "08": ("AGO", 8),
    "09": ("SET", 9), "10": ("OUT", 10), "11": ("NOV", 11), "12": ("DEZ", 12),
}



def clean_sku(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    return s if s else np.nan


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)


def extrair_produto_por_sku(df, col_sku, col_produto):
    aux = (
        df[[col_sku, col_produto]]
        .dropna(subset=[col_sku, col_produto])
        .astype(str)
    )

    produto_sku = (
        aux.groupby(col_sku)[col_produto]
        .agg(lambda x: x.mode().iloc[0] if not x.mode().empty else x.iloc[-1])
        .reset_index()
        .rename(columns={
            col_sku: "SKU",
            col_produto: "produto"
        })
    )

    return produto_sku


def processar_collection(collection_path):
    df = pd.read_excel(collection_path, engine="openpyxl")

    
    df["SKU"] = df["SKU do produto (seller_custom_field)"].apply(clean_sku)

  
    if COL_PRODUTO not in df.columns:
        raise ValueError(f"❌ Coluna de produto não encontrada: {COL_PRODUTO}")

    produto_sku = extrair_produto_por_sku(
        df,
        col_sku="SKU",
        col_produto=COL_PRODUTO
    )


    df["net_received_amount"] = pd.to_numeric(
        df["Valor total recebido (net_received_amount)"], errors="coerce"
    ).fillna(0)

    df["transaction_amount"] = pd.to_numeric(
        df["Valor do produto (transaction_amount)"], errors="coerce"
    ).fillna(0)

 
    df["qtd"] = np.where(
        df["transaction_amount"] > 0, 1,
        np.where(df["transaction_amount"] < 0, -1, 0)
    )

    sku_qty = (
        df.dropna(subset=["SKU"])
          .groupby("SKU", as_index=False)["qtd"]
          .sum()
          .query("qtd != 0")
    )


    receita_sku = (
        df.dropna(subset=["SKU"])
          .groupby("SKU", as_index=False)["net_received_amount"]
          .sum()
          .rename(columns={"net_received_amount": "receita_liquida_sku"})
    )


    faturamento_sku = (
        df.dropna(subset=["SKU"])
          .groupby("SKU", as_index=False)["transaction_amount"]
          .sum()
          .rename(columns={"transaction_amount": "faturamento_bruto_sku"})
    )

  
    dimi = pd.read_excel(DIMI_PATH, sheet_name=DIMI_SHEET)
    dimi = dimi.rename(columns={dimi.columns[0]: "SKU"})
    dimi["SKU"] = dimi["SKU"].apply(clean_sku)

    custo_col = [c for c in dimi.columns if c.lower() == "custo"][0]
    dimi[custo_col] = pd.to_numeric(dimi[custo_col], errors="coerce")

    dimi_cost = (
        dimi[["SKU", custo_col]]
        .dropna(subset=["SKU"])
        .drop_duplicates(subset=["SKU"], keep="last")
        .rename(columns={custo_col: "custo_unit"})
    )


    resultado = (
        sku_qty
        .merge(receita_sku, on="SKU", how="left")
        .merge(faturamento_sku, on="SKU", how="left")  
        .merge(dimi_cost, on="SKU", how="left")
        .merge(produto_sku, on="SKU", how="left")
    )

    resultado["cmv_total_sku"] = resultado["qtd"] * resultado["custo_unit"]
    resultado["lucro_sku"] = resultado["receita_liquida_sku"] - resultado["cmv_total_sku"]
    resultado["margem_sku"] = resultado["lucro_sku"] / resultado["receita_liquida_sku"]

    return resultado



ensure_dir(OUTDIR)

for pasta_mes in sorted(os.listdir(BASE_MESES)):
    if pasta_mes not in MAPA_MESES:
        continue

    mes_nome, mes_num = MAPA_MESES[pasta_mes]
    pasta_completa = os.path.join(BASE_MESES, pasta_mes)

    arquivos = [f for f in os.listdir(pasta_completa) if f.endswith(".xlsx")]
    if not arquivos:
        print(f"⚠️ Nenhum arquivo em {pasta_mes}")
        continue

    collection_path = os.path.join(pasta_completa, arquivos[0])
    print(f"📊 Processando {mes_nome}/{ANO}: {collection_path}")

    resultado_sku = processar_collection(collection_path)

    nome_saida = f"DRE_ML_{mes_nome}{str(ANO)[-2:]}.xlsx"
    out_path = os.path.join(OUTDIR, nome_saida)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        resultado_sku.to_excel(writer, index=False, sheet_name="Lucro_por_SKU")

    print(f"✅ Gerado: {nome_saida}")

print("\n🎉 Todos os meses processados com sucesso.")
