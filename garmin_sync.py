import pandas as pd
from garminconnect import Garmin
import my_garmin_common
from dotenv import load_dotenv
from datetime import datetime
from zoneinfo import ZoneInfo
import os
import pprint

# .envファイルから環境変数を読み込む
load_dotenv()

# --- 設定項目 ---
GARMIN_EMAIL = my_garmin_common.get_secret("GARMIN_EMAIL")
GARMIN_PASSWORD = my_garmin_common.get_secret("GARMIN_PASSWORD")
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
        # アジア/東京のタイムゾーンで取得
        now_jst = datetime.now(ZoneInfo("Asia/Tokyo"))

        date_str = now_jst.strftime("%Y-%m-%d %H:%M:%S")
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
            'body_battery': stats.get('bodyBatteryMostRecentValue'), # 最新のBodyBattery
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


def main():
    """メイン処理"""
    # 1. スプレッドシート接続（先に確認）
    if not SA_KEY_VALUE:
        print("❌ シークレット 'SA_KEY' が設定されていません。")
        return

    spreadsheet = my_garmin_common.get_spreadsheet(SA_KEY_VALUE)
    if not spreadsheet:
        print("❌ スプレッドシートに接続できないため終了します。")
        return

    # 2. Garminログイン
    garmin = login_to_garmin()
    if not garmin:
        return

    # 3. データの取得
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

    # 4. マージと保存（daily_summaryへ上書き）
    # 最新の日付から順に処理
    for data in new_data_list:
         my_garmin_common.update_daily_summary(spreadsheet, data)

    # 5. ログ用シート(Sheet1)への追記（最新データのみ）
    if new_data_list:
        latest_data = new_data_list[0] # リストの先頭が最新（今日）
        my_garmin_common.append_to_log(spreadsheet, latest_data)

if __name__ == "__main__":
    main()
