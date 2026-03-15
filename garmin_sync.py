import pandas as pd
from garminconnect import Garmin
import my_garmin_common
from dotenv import load_dotenv
import datetime
import os

# .envファイルから環境変数を読み込む
load_dotenv()

# --- 設定項目 ---
GARMIN_EMAIL = my_garmin_common.get_secret("GARMIN_EMAIL")
GARMIN_PASSWORD = my_garmin_common.get_secret("GARMIN_PASSWORD")
SPREADSHEET_ID = '15CCDjcBCqSWYacPWf_RNXTBdJZ33x6pXAc1PhwPfkiY'
SA_KEY_VALUE = my_garmin_common.get_secret("SA_KEY")


def login_to_garmin():
    """Garmin Connectにログインする"""
    if not GARMIN_EMAIL or not GARMIN_PASSWORD:
        print("❌ エラー: 環境変数 'GARMIN_EMAIL' と 'GARMIN_PASSWORD' が設定されていません。")
        return None
    
    try:
        garmin = Garmin(GARMIN_EMAIL, GARMIN_PASSWORD)
        garmin.login()
        print("✅ Garmin login successful.")
        return garmin
    except Exception as e:
        print(f"❌ Garmin Login Failed: {e}")
        return None


def fetch_daily_data(garmin_client, date_obj):
    """指定日のヘルスケアデータを取得して辞書で返す"""
    try:
        date_str = date_obj.strftime("%Y-%m-%d")
        stats = garmin_client.get_user_summary(date_str)
        
        # 必要なデータを抽出（項目は必要に応じて調整してください）
        # 単位変換とデータ加工
        # 睡眠: 秒 -> 時間
        sleep_seconds = stats.get('sleepingSeconds')
        sleep_hours = round(sleep_seconds / 3600, 2) if sleep_seconds else None

        # 体重: グラム -> キログラム (Garmin APIはグラムで返すことがあるため念のため)
        weight = stats.get('totalWeight')
        if weight and weight > 1000:
            weight = round(weight / 1000, 2)

        # スプレッドシートの列順序に合わせて辞書を作成
        data_dict = {
            'calendarDate': date_str,
            'steps': stats.get('totalSteps'),
            'heart_rate': stats.get('restingHeartRate'),
            'stress': stats.get('averageStressLevel'),
            'body_battery': None, # Summaryからは直接取れないため一旦None
            'sleep_hours': sleep_hours,
            'weight': weight,
            'bmi': stats.get('bodyMassIndex'),
            'body_fat_pct': stats.get('bodyFat'),
            'muscle_pct': stats.get('muscleMass'),
            'visceral_fat': stats.get('visceralFat'),
            'metabolism': stats.get('bmr'),
            'bone_mass': stats.get('boneMass'),
        }
        print(f"✅ Data fetched for {date_str}")
        print(f"   📊 {data_dict}")
        return data_dict
    except Exception as e:
        print(f"⚠️ Failed to fetch data for {date_obj}: {e}")
        return None


def sync_to_google_sheets(data_dict):
    """Googleスプレッドシートにデータを同期する"""
    print("\n--- Google Sheetsへの同期を開始 ---")
    
    if not SA_KEY_VALUE:
        print(f"⚠️ 警告: シークレット 'SA_KEY' が設定されていません。")
        print("   Googleスプレッドシートへの同期をスキップします。")
        return

    client = my_garmin_common.get_google_creds(SA_KEY_VALUE)
    if not client:
        return

    # サービスアカウントのメールアドレスを表示しておくと便利
    try:
        print(f"ℹ️  Service Account Email: {client.auth.service_account_email}")
    except:
        pass

    try:
        sheet = client.open_by_key(SPREADSHEET_ID).sheet1
        # データの追加ロジック（例：末尾に追加）
        # 実際には日付の重複チェックなどが必要かもしれませんが、ここでは単純追加の例
        values = list(data_dict.values())
        sheet.append_row(values)
        print("✅ Google Sheets updated.")
    except Exception as e:
        print(f"❌ Google Sheets Sync Error: {e}")
        if "403" in str(e):
            print("💡 ヒント: 上記のサービスアカウントのメールアドレスをコピーし、")
            print("   スプレッドシートの「共有」ボタンから「編集者」として追加してください。")


def main():
    """メイン処理"""
    # 1. Garminログイン
    garmin = login_to_garmin()
    if not garmin:
        return

    # 2. 既存データをCSVから読み込み
    df_current = my_garmin_common.load_data()
    print(f"Current data rows: {len(df_current)}")

    # 3. データの取得（例：今日と昨日のデータを確認）
    today = datetime.date.today()
    # 必要に応じて遡る日数を変えてください
    target_dates = [today - datetime.timedelta(days=i) for i in range(2)] 
    
    new_data_list = []
    
    for date_obj in target_dates:
        data = fetch_daily_data(garmin, date_obj)
        if data:
            new_data_list.append(data)

    if not new_data_list:
        print("No new data fetched.")
        return

    # 4. CSVの更新と保存
    df_new = pd.DataFrame(new_data_list)
    df_new['calendarDate'] = df_new['calendarDate'].astype(str)
    df_new = df_new.set_index('calendarDate')
    
    # 既存データと結合（新しいデータを優先して上書き）
    df_updated = df_new.combine_first(df_current)
    my_garmin_common.save_data(df_updated)

    # 5. 最新データ（今日分）だけスプレッドシートに送る例
    # 全データを同期するとAPI制限にかかりやすいため、最新1件のみなどを推奨
    if new_data_list:
        latest_data = new_data_list[0] # リストの先頭が最新（今日）
        sync_to_google_sheets(latest_data)

if __name__ == "__main__":
    main()
