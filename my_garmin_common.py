import pandas as pd
import os

CSV_PATH = '/content/drive/MyDrive/Colab Notebooks/sakura_body.csv'

def load_data():
    # ファイルが見つからない場合、Colab環境であればGoogleドライブのマウントを試みる
    if not os.path.exists(CSV_PATH):
        try:
            from google.colab import drive
            drive.mount('/content/drive')
        except ImportError:
            pass

    if os.path.exists(CSV_PATH):
        df = pd.read_csv(CSV_PATH)
        df['calendarDate'] = df['calendarDate'].astype(str)
        return df.set_index('calendarDate')
    return pd.DataFrame()

def save_data(df):
    df = df[~df.index.duplicated(keep='last')]
    df = df.sort_index(ascending=False)
    df.to_csv(CSV_PATH, index_label='calendarDate')
    print('✨ CSV updated: ' + CSV_PATH)

load_data()