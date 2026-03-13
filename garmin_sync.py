import pandas as pd
from garminconnect import Garmin
import my_garmin_common
import traceback
from dotenv import load_dotenv

# .envファイルから環境変数を読み込む
load_dotenv()

# --- 設定項目 ---
# 環境変数から認証情報を読み込む
GARMIN_EMAIL = my_garmin_common.get_secret("GARMIN_EMAIL")
GARMIN_PASSWORD = my_garmin_common.get_secret("GARMIN_PASSWORD")

# GoogleスプレッドシートのID
SPREADSHEET_ID = '15CCDjcBCqSWYacPWf_RNXTBdJZ33x6pXAc1PhwPfkiY'
# Googleサービスアカウントの秘密鍵(JSONファイル)へのパス
GOOGLE_SECRETS_PATH = my_garmin_common.get_secret("GOOGLE_SECRETS_PATH")


def main():
    """メイン処理"""
    # 認証情報が設定されているか確認
    if not GARMIN_EMAIL or not GARMIN_PASSWORD:
        print("❌ エラー: 環境変数 'GARMIN_EMAIL' と 'GARMIN_PASSWORD' が設定されていません。")
        print("   .envファイルを作成するか、環境変数を設定してください。")
        return

    # 1. 既存データをCSVから読み込み
    df_current = my_garmin_common.load_data()
    

        # 5. Googleスプレッドシートへ同期
        print("\n--- Google Sheetsへの同期を開始 ---")
        if not GOOGLE_SECRETS_PATH or not os.path.exists(GOOGLE_SECRETS_PATH):
            print(f"⚠️  警告: 環境変数 'GOOGLE_SECRETS_PATH' が設定されていないか、ファイルが見つかりません。")
            print(f"     指定されたパス: {GOOGLE_SECRETS_PATH}")
            print("     Googleスプレッドシートへの同期をスキップします。")
        else:
            df_for_sheet = pd.DataFrame([data_dict])
            google_creds = my_garmin_common.get_google_creds(GOOGLE_SECRETS_PATH)
