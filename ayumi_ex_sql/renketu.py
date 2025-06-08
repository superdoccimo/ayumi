import os
import pandas as pd
from openpyxl import Workbook

import config

# Excelファイルのパスを指定
input_excel_path = config.EXCEL_OUTPUT_PATH
output_excel_path = config.CONSOLIDATED_OUTPUT_PATH

# 新しいワークブックとシートを作成
wb = Workbook()
ws_consolidated = wb.active
ws_consolidated.title = "Consolidated"

# パンダスでExcelファイルを読み込む
if not os.path.exists(input_excel_path):
    raise FileNotFoundError(f"{input_excel_path} が見つかりません")

with pd.ExcelFile(input_excel_path) as xls:
    first_sheet_processed = False
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        if not first_sheet_processed:
            ws_consolidated.append(df.columns.tolist())
            first_sheet_processed = True
        for row in df.itertuples(index=False):
            ws_consolidated.append(row)

# 新しいExcelファイルに保存
wb.save(output_excel_path)
print(f'データが結合され、{output_excel_path} に保存されました。')
