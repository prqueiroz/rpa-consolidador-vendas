import pandas as pd
import logging
from src import config

def gerar_relatorio_excel_com_dashboard(df: pd.DataFrame):
    caminho_saida = config.FINAL_REPORT_PATH
    print(" Gerando dashboard e formatando Excel...")

    with pd.ExcelWriter(caminho_saida, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Vendas_Consolidadas", index=False)
        
        workbook = writer.book
        worksheet_dados = writer.sheets["Vendas_Consolidadas"]

        header_format = workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "vcenter",
            "align": "center",
            "fg_color": "#1F4E78",
            "font_color": "#FFFFFF",
            "border": 1
        })
        
        currency_format = workbook.add_format({
            "num_format": "R$ #,##0.00",
            "align": "right"
        })
        
        date_format = workbook.add_format({
            "num_format": "yyyy-mm-dd",
            "align": "center"
        })

       
        for col_num, value in enumerate(df.columns.values):
            worksheet_dados.write(0, col_num, value, header_format)
            worksheet_dados.set_column(col_num, col_num, 20)

        
        worksheet_dados.set_column("A:A", 14, date_format)
        worksheet_dados.set_column("E:F", 16, currency_format)
        worksheet_dados.set_column("I:I", 16, currency_format)

      
        df_resumo = df.groupby("Filial")["Valor_Liquido"].sum().reset_index()
        df_resumo.columns = ["Filial", "Faturamento_Total"]
        df_resumo.to_excel(writer, sheet_name="Dashboard", index=False)

        worksheet_dash = writer.sheets["Dashboard"]

        for col_num, value in enumerate(df_resumo.columns.values):
            worksheet_dash.write(0, col_num, value, header_format)
            worksheet_dash.set_column(col_num, col_num, 20)

        worksheet_dash.set_column("B:B", 20, currency_format)

        
        chart = workbook.add_chart({"type": "column"})
        qtd_filiais = len(df_resumo)
        
        chart.add_series({
            "name": "=Dashboard!$B$1",
            "categories": f"=Dashboard!$A$2:$A${qtd_filiais + 1}",
            "values": f"=Dashboard!$B$2:$B${qtd_filiais + 1}",
            "fill": {"color": "#2F5597"},
            "data_labels": {"value": True, "num_format": "R$ #,##0"}
        })

        chart.set_title({"name": "Faturamento Total por Filial (R$)"})
        chart.set_x_axis({"name": "Filial"})
        chart.set_y_axis({"name": "Valor Total (R$)"})
        chart.set_legend({"position": "none"})

        worksheet_dash.insert_chart("D2", chart)

    print(f" Relatório atualizado com sucesso em: {caminho_saida}")