import sys
import logging
from src import config, etl, dashboard

def orquestrar_automacao():
    """Função principal que gerencia o fluxo de execução do RPA."""
    print("=" * 60)
    print("🚀 INICIANDO PROCESSAMENTO RPA - CONSOLIDAÇÃO DE VENDAS")
    print("=" * 60)
    
    logging.info("Iniciando execução da automação RPA.")

    try:
        
        df_consolidado = etl.processar_e_consolidar_dados()

        
        dashboard.gerar_relatorio_excel_com_dashboard(df_consolidado)

        print("\n✨ AUTOMATION COMPLETED WITH SUCCESS!")
        print(f"📁 Verifique o arquivo final em: {config.FINAL_REPORT_PATH}")
        print("=" * 60)
        
    except Exception as e:
        erro_msg = f"Falha crítica durante a execução do robô: {str(e)}"
        print(f"\n❌ ERRO: {erro_msg}")
        logging.critical(erro_msg, exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    orquestrar_automacao()