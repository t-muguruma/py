import pandas as pd
import os
import json
from google.oauth2.service_account import Credentials
import gspread

def get_secret(key):
    """Colabのシークレット、または環境変数から値を取得する。Colabを優先する。"""
    # 1. Colab環境のシークレットを確認
    try:
        from google.colab import userdata  # type: ignore
        value = userdata.get(key)
        if value is not None:
            return value
    except ImportError:
        # Colabライブラリがない場合はローカル環境と判断し、次に進む
        pass
    
    # 2. 環境変数（ローカルの.envなど）を確認
    return os.environ.get(key)

def get_google_creds(secret_value):
    """
    Googleサービスアカウントの認証を行い、クライアントを返す。
    引数がファイルパスでも、JSON文字列でも対応する。
    """
    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    try:
        # Case 1: 引数がファイルパスの場合
        if os.path.exists(str(secret_value)):
            creds = Credentials.from_service_account_file(secret_value, scopes=scope)
            return gspread.authorize(creds)
        # Case 2: 引数がJSON文字列の場合
        else:
            info = json.loads(secret_value)
            creds = Credentials.from_service_account_info(info, scopes=scope)
            return gspread.authorize(creds)
    except Exception as e:
        print(f"⚠️ Google Auth Error: {e}")
        return None

def sheet_to_df(sheet):
    """スプレッドシートのデータを読み込み、DataFrameにして返す"""
    data = sheet.get_all_records()
    if not data:
        return pd.DataFrame()
    df = pd.DataFrame(data)
    # calendarDateを文字列にし、インデックスに設定
    if 'calendarDate' in df.columns:
        df['calendarDate'] = df['calendarDate'].astype(str)
        df = df.set_index('calendarDate')
    return df

def df_to_sheet(sheet, df):
    """DataFrameをスプレッドシートに全上書き保存する"""
    # NaN（欠損値）を空文字に置換（JSON化エラー防止）
    df = df.fillna('')
    # インデックス(calendarDate)を列に戻す
    df_reset = df.reset_index()
    # ヘッダーとデータをリスト化
    data = [df_reset.columns.values.tolist()] + df_reset.values.tolist()
    
    sheet.clear()
    sheet.update('A1', data)
    print(f"✨ Sheet '{sheet.title}' updated.")
