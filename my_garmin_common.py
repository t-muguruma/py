import pandas as pd
import os
import json
from google.oauth2.service_account import Credentials
import gspread

CSV_PATH = '/content/drive/MyDrive/Colab Notebooks/sakura_body.csv'

def mount_drive():
    """Colab環境でGoogleドライブをマウントする"""
    try:
        from google.colab import drive  # type: ignore
        # まだマウントされていない場合のみ実行
        if not os.path.exists('/content/drive'):
            drive.mount('/content/drive')
    except ImportError:
        pass

def load_data():
    # データ読み込み前にマウントを試みる
    if not os.path.exists(CSV_PATH):
        mount_drive()

    if os.path.exists(CSV_PATH):
        df = pd.read_csv(CSV_PATH)
        df['calendarDate'] = df['calendarDate'].astype(str)
        return df.set_index('calendarDate')
    return pd.DataFrame()

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

def save_data(df):
    df = df[~df.index.duplicated(keep='last')]
    df = df.sort_index(ascending=False)
    df.to_csv(CSV_PATH, index_label='calendarDate')
    print('✨ CSV updated: ' + CSV_PATH)

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
