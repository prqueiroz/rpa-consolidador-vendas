import os
from pathlib import Path

# Caminho base do projeto (raiz do repositório)
BASE_DIR = Path(__file__).resolve().parent.parent

# Diretores de Dados
DATA_DIR = BASE_DIR / "data"
INPUT_DIR = DATA_DIR / "input"
OUTPUT_DIR = DATA_DIR / "output"

# Diretorio de Logs
LOGS_DIR = BASE_DIR / "logs"

# Garantir que as pastas de saída e log existam no sistema
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOGS_DIR.mkdir(parents=True, exist_ok=True)

# Caminho do relatório de saída
FINAL_REPORT_PATH = OUTPUT_DIR / "Relatorio_Consolidado_Vendas.xlsx"

# Nomes dos arquivos de entrada esperados
FILE_SP = INPUT_DIR / "filial_sp.csv"
FILE_MG = INPUT_DIR / "filial_mg.xlsx"
FILE_RJ = INPUT_DIR / "filial_rj.xlsx"