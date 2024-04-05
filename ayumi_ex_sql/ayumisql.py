import pandas as pd
import hashlib
import os
import json
from datetime import datetime
import time
import mysql.connector

def generate_hash(row):
    """行の内容からハッシュ値を生成"""
    row_str = ''.join(map(str, row.values))
    return hashlib.sha256(row_str.encode()).hexdigest()

def check_for_new_data():
    # JSONファイルのパス（以前のデータを保持）
    json_file_path = 'c:/ayumi/data.json'
    
    # Excelファイルとシートのパス（読み込み用）
    file_path = 'c:/ayumi/ayumi.xlsm'
    sheet_name = 'Sheet2'
    
    # Excelファイルからデータを読み込む
    while True:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            break
        except PermissionError:
            print("ファイルにアクセスできません。2秒後に再試行します...")
            time.sleep(2)
    
    # 最初の300行を選択（2行目から301行目に相当）
    df = df.iloc[0:300]
    
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
        print("新しいデータが検出されました。以下のデータが追加されました:")
        print(df_new_unique.drop(columns=['hash']).to_string(index=False))
        
        # MySQLデータベースに接続
        cnx = mysql.connector.connect(
            host='192.168.0.139',  # ホスト名
            database='test',  # データベース名
            user='root',  # ユーザー名
            password='testpass'  # パスワード
        )
        cursor = cnx.cursor()
        
        # テーブルが存在しない場合は作成
        create_table_query = """
        CREATE TABLE IF NOT EXISTS trade_data (
            id INT AUTO_INCREMENT PRIMARY KEY,
            trade_time DATETIME,
            volume INT,
            price FLOAT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
        """
        cursor.execute(create_table_query)
        
        # 新しいデータを逆順に並べ替えてMySQLデータベースに挿入
        for _, row in df_new_unique.drop(columns=['hash']).iloc[::-1].iterrows():
            if '--------' not in str(row['時刻']):  # 無効な日時値をフィルタリング
                current_date = datetime.now().strftime('%Y-%m-%d')  # 現在の日付を取得
                trade_time = f"{current_date} {row['時刻']}"  # 日付と時刻を結合
                insert_query = """
                INSERT INTO trade_data (trade_time, volume, price)
                VALUES (%s, %s, %s)
                """
                values = (trade_time, row['出来高'], row['約定値'])
                cursor.execute(insert_query, values)
        
        cnx.commit()
        cursor.close()
        cnx.close()
        
        # 全データ（新旧）を統合してJSONファイルに保存
        df_full = pd.concat([df_previous, df_new_unique[['hash']]])
        df_full.to_json(json_file_path, orient='records', force_ascii=False)
        print("データの処理が完了しました。")
    else:
        print("新しいデータは検出されませんでした。")

# 定期的にデータチェックを行うループ
while True:
    check_for_new_data()
    time.sleep(5)  # 5秒待機