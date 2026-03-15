import pandas as pd
from garminconnect import Garmin
import my_garmin_common
from dotenv import load_dotenv
import datetime
import os
import pprint

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

        # 🔍 デバッグ用: APIから返ってきた生の全データを表示
        print(f"--- Raw Data for {date_str} ---")
        pprint.pprint(stats)
        
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
            # --- 活動量 ---
            'steps': stats.get('totalSteps'),
            'distance_m': stats.get('totalDistanceMeters'), # 移動距離(m)
            'floors_ascended': stats.get('floorsAscended'), # 上った階数
            'active_calories': stats.get('activeKilocalories'), # 活動カロリー
            'total_calories': stats.get('totalKilocalories'), # 総消費カロリー
            # --- 心拍・ストレス ---
            'heart_rate': stats.get('restingHeartRate'),
            'max_heart_rate': stats.get('maxHeartRate'), # 最大心拍
            'min_heart_rate': stats.get('minHeartRate'), # 最小心拍
            'stress': stats.get('averageStressLevel'), 
            'body_battery': stats.get('maxBodyBattery'), # 取れるか不明だが一応追加
            # --- 睡眠・回復 ---
            'sleep_hours': sleep_hours,
            # --- 体組成 ---
            'weight': weight,
            'bmi': stats.get('bodyMassIndex'),
            'body_fat_pct': stats.get('bodyFat'),
            'muscle_pct': stats.get('muscleMass'),
            'visceral_fat': stats.get('visceralFat'),
            'metabolism': stats.get('bmr'),
            'bone_mass': stats.get('boneMass'),
            'water_ml': stats.get('totalWaterIntake'), # 水分摂取
            # --- 運動時間 ---
            'moderate_minutes': stats.get('moderateIntensityMinutes'), # 中強度運動(分)
            'vigorous_minutes': stats.get('vigorousIntensityMinutes'), # 高強度運動(分)
        }
        print(f"✅ Data fetched for {date_str}")
        print(f"   📊 {data_dict}")
        return data_dict
    except Exception as e:
        print(f"⚠️ Failed to fetch data for {date_obj}: {e}")
        return None


def get_spreadsheet_client():
    """スプレッドシートクライアントとシートオブジェクトを取得"""
    if not SA_KEY_VALUE:
        print(f"⚠️ 警告: シークレット 'SA_KEY' が設定されていません。")
        return None, None

    client = my_garmin_common.get_google_creds(SA_KEY_VALUE)
    if not client:
        return None, None

    try:
        # スプレッドシートを開く
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        
        # サービスアカウントのメールアドレスを表示（権限確認用）
        try:
            print(f"ℹ️  Service Account Email: {client.auth.service_account_email}")
        except:
            pass
            
        return client, spreadsheet
    except Exception as e:
        print(f"❌ Spreadsheet Connection Error: {e}")
        if "403" in str(e):
            print("💡 ヒント: 上記のサービスアカウントのメールアドレスをコピーし、")
            print("   スプレッドシートの「共有」ボタンから「編集者」として追加してください。")
        return None, None


def main():
    """メイン処理"""
    # 1. スプレッドシート接続（先に確認）
    client, spreadsheet = get_spreadsheet_client()
    if not spreadsheet:
        print("❌ スプレッドシートに接続できないため終了します。")
        return

    # シートの準備
    # 1) sheet1: ログ用（単純追記）
    log_sheet = spreadsheet.sheet1
    
    # 2) daily_summary: マスタ用（CSVの代わり）
    SUMMARY_SHEET_NAME = 'daily_summary'
    try:
        summary_sheet = spreadsheet.worksheet(SUMMARY_SHEET_NAME)
    except:
        print(f"Creating new sheet: {SUMMARY_SHEET_NAME}")
        summary_sheet = spreadsheet.add_worksheet(title=SUMMARY_SHEET_NAME, rows=1000, cols=20)

    # 2. Garminログイン
    garmin = login_to_garmin()
    if not garmin:
        return

    # 3. 既存データをスプレッドシート(daily_summary)から読み込み
    print("Loading existing data from sheet...")
    df_current = my_garmin_common.sheet_to_df(summary_sheet)
    print(f"Current data rows: {len(df_current)}")

    # 4. データの取得
    today = datetime.date.today()
    target_dates = [today - datetime.timedelta(days=i) for i in range(2)] 
    
    new_data_list = []
    for date_obj in target_dates:
        data = fetch_daily_data(garmin, date_obj)
        if data:
            new_data_list.append(data)

    if not new_data_list:
        print("No new data fetched.")
        return

    # 5. マージと保存（daily_summaryへ上書き）
    df_new = pd.DataFrame(new_data_list)
    df_new['calendarDate'] = df_new['calendarDate'].astype(str)
    df_new = df_new.set_index('calendarDate')
    
    df_updated = df_new.combine_first(df_current)
    
    # ソートして保存
    df_updated = df_updated.sort_index(ascending=False)
    my_garmin_common.df_to_sheet(summary_sheet, df_updated)

    # 6. ログ用シート(Sheet1)への追記（最新データのみ）
    if new_data_list:
        print("\n--- Appending to Log Sheet ---")
        latest_data = new_data_list[0] # リストの先頭が最新（今日）
        
        # ログシート用に日時スタンプを作成し、先頭の年月日を置き換える
        log_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        values = list(latest_data.values())
        values[0] = log_timestamp
        
        log_sheet.append_row(values)
        print("✅ Log sheet updated.")

if __name__ == "__main__":
    main()
