import json
import os
import time
from datetime import datetime
import pandas as pd
import hashlib
import traceback

import config

def generate_hash(row):
    """行の内容からハッシュ値を生成"""
    row_str = ''.join(map(str, row.values))
    return hashlib.sha256(row_str.encode()).hexdigest()

def load_last_row(progress_path):
    if os.path.exists(progress_path):
        with open(progress_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get("last_row", 0)
    return 0


def save_last_row(progress_path, row):
    with open(progress_path, "w", encoding="utf-8") as f:
        json.dump({"last_row": row}, f)


def check_for_new_data():
    json_file_path = config.JSON_PATH
    file_path = config.EXCEL_PATH
    sheet_name = config.SHEET_NAME
    progress_path = config.PROGRESS_PATH
    
    # Excelファイルからデータを読み込む
    while True:
        try:
            df_all = pd.read_excel(file_path, sheet_name=sheet_name)
            break
        except PermissionError:
            print("ファイルにアクセスできません。5秒後に再試行します...")
            time.sleep(5)

    last_row = load_last_row(progress_path)
    df = df_all.iloc[last_row:]
    
    # ハッシュ値を計算して新しい列に追加
    df['hash'] = df.apply(generate_hash, axis=1)
    
    # 既存のJSONファイルからデータを読み込む
    if os.path.exists(json_file_path):
        with open(json_file_path, 'r', encoding='utf-8') as file:
            previous_data = json.load(file)
        df_previous = pd.DataFrame(previous_data)
        if 'hash' not in df_previous.columns:
            df_previous['hash'] = [None] * len(df_previous)  # 'hash'列がない場合は追加
    else:
        df_previous = pd.DataFrame(columns=['hash'])  # 空のDataFrameに'hash'列を設定
    
    # 新しいデータのハッシュ値が以前のデータに存在しない場合、そのデータを新規追加として扱う
    df_new_unique = df[~df['hash'].isin(df_previous['hash'])]
    
    # 新規データがあれば、それを処理
    if not df_new_unique.empty:
        print("新しいユニークなデータが検出されました。以下のデータが追加されました:")
        print(df_new_unique.drop(columns=['hash']).to_string(index=False))
        
        excel_output_path = config.EXCEL_OUTPUT_PATH
        
        # 新しいシート名（例：現在の日時を使用）
        new_sheet_name = datetime.now().strftime('Data_%Y-%m-%d_%H-%M-%S')
        
        # 無効な日時値をフィルタリング
        df_new_unique = df_new_unique[~df_new_unique['時刻'].astype(str).str.contains('--------')]
        
        # エクセルファイルに新しいシートを追加
        with pd.ExcelWriter(excel_output_path, engine='openpyxl', mode='a' if os.path.exists(excel_output_path) else 'w') as writer:
            df_new_unique.drop(columns=['hash']).iloc[::-1].to_excel(writer, sheet_name=new_sheet_name, index=False)
        
        df_full = pd.concat([df_previous, df_new_unique[['hash']]])
        df_full.to_json(json_file_path, orient='records', force_ascii=False)
        save_last_row(progress_path, len(df_all))
        print("データの処理が完了しました。")
    else:
        print("新しいユニークなデータは検出されませんでした。")
        save_last_row(progress_path, len(df_all))

# 定期的にデータチェックを行うループ
while True:
    try:
        check_for_new_data()
    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
    time.sleep(5)