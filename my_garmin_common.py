import pandas as pd
import os

CSV_PATH = '/content/drive/MyDrive/Colab Notebooks/sakura_body.csv'

def mount_drive():
    """Colab環境でGoogleドライブをマウントする"""
    try:
        from google.colab import drive
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
        from google.colab import userdata
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
