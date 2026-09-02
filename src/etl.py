import pandas as pd
import logging
from src import config

logging.basicConfig(
    filename=config.LOGS_DIR / "execucao.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    encoding="utf-8"
)

def limpar_valor_financeiro(valor) -> float:
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    
    str_val = str(valor).strip().replace("R$", "").replace(" ", "")
    if "," in str_val:
        str_val = str_val.replace(".", "").replace(",", ".")
    return float(str_val)

def processar_e_consolidar_dados() -> pd.DataFrame:
    lista_dfs = []
    
    print("🤖 Iniciando extração e limpeza dos dados...")

   
    if config.FILE_SP.exists():
        df_sp = pd.read_csv(config.FILE_SP, encoding="utf-8-sig")
        df_sp["Filial"] = "São Paulo"
        lista_dfs.append(df_sp)

   
    if config.FILE_MG.exists():
        df_mg = pd.read_excel(config.FILE_MG)
        df_mg["Filial"] = "Minas Gerais"
        lista_dfs.append(df_mg)

    
    if config.FILE_RJ.exists():
        df_rj = pd.read_excel(config.FILE_RJ)
        df_rj["Filial"] = "Rio de Janeiro"
        lista_dfs.append(df_rj)

    df_consolidado = pd.concat(lista_dfs, ignore_index=True)

    
    colunas_texto = ["Codigo_Cliente", "Nome_Cliente", "Categoria_Produto", "Canal_Venda"]
    for col in colunas_texto:
        df_consolidado[col] = df_consolidado[col].astype(str).str.strip()

  
    df_consolidado["Valor_Bruto"] = df_consolidado["Valor_Bruto"].apply(limpar_valor_financeiro)
    df_consolidado["Valor_Desconto"] = df_consolidado["Valor_Desconto"].apply(limpar_valor_financeiro)
    df_consolidado["Valor_Liquido"] = df_consolidado["Valor_Bruto"] - df_consolidado["Valor_Desconto"]

    
    df_consolidado["Data_Venda"] = pd.to_datetime(
        df_consolidado["Data_Venda"].astype(str).str.strip(),
        format="mixed",
        yearfirst=False
    ).dt.strftime("%Y-%m-%d")


    df_consolidado.drop_duplicates(
        subset=["Data_Venda", "Codigo_Cliente", "Valor_Bruto"],
        keep="first",
        inplace=True
    )

    print(f"✅ Dados consolidados! Total de registros limpos: {len(df_consolidado)}")
    return df_consolidado