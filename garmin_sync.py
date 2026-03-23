import garth
import pandas as pd
from garminconnect import Garmin
import my_garmin_common
from dotenv import load_dotenv
import datetime
from zoneinfo import ZoneInfo
import os
import sys
import time

# ログ一貫性のためのフォーマット
def log_message(level, message):
    levels = {"INFO": "✅", "ERROR": "❌", "WARN": "⚠️", "DEBUG": "🔍"}
    prefix = levels.get(level, "ℹ️")
    print(f"{prefix} {message}")

# --- 環境設定 ---
load_dotenv()
GARMIN_EMAIL = my_garmin_common.get_secret("GARMIN_EMAIL")
GARMIN_PASSWORD = my_garmin_common.get_secret("GARMIN_PASSWORD")
SA_KEY_VALUE = my_garmin_common.get_secret("SA_KEY")

# Garmin認証クライアント
def get_garmin_client():
    token_dir = "./.garth"
    
    # 1. キャッシュ（合鍵）があればそれを使う
    if os.path.exists(token_dir):
        try:
            garth.resume(token_dir)
            garmin = Garmin(GARMIN_EMAIL, GARMIN_PASSWORD)
            garmin.garth = garth.client

            # 合鍵からユーザー名を復元してセットする（これが無いと daily/None になってエラーになる）
            if garth.client.profile and "displayName" in garth.client.profile:
                garmin.display_name = garth.client.profile["displayName"]
                log_message("INFO", f"✅ キャッシュからセッションを復元成功。ユーザー: {garmin.display_name}")
                return garmin
            
            raise Exception("キャッシュにユーザー情報(displayName)が見つかりません。")
        except Exception as e:
            log_message("WARN", f"キャッシュからの復元に失敗しました。通常ログインを試みます: {e}")

    # 2. キャッシュがない、または失敗した場合は通常ログイン
    if not GARMIN_EMAIL or not GARMIN_PASSWORD:
        log_message("ERROR", "環境変数 'GARMIN_EMAIL' または 'GARMIN_PASSWORD' が未設定。")
        return None
    try:
        garmin = Garmin(GARMIN_EMAIL, GARMIN_PASSWORD)
        garmin.login()
        log_message("INFO", "Garmin Connectへのログイン成功。")
        return garmin
    except Exception as e:
        log_message("ERROR", f"Garminへのログイン失敗: {e}")
        return None

# データ取得用共通関数
def fetch_from_garmin(api_function, date, *args, **kwargs):
    try:
        result = api_function(*args, **kwargs)
        log_message("INFO", f"{date}: データ取得成功。")
        return result
    except Exception as e:
        log_message("WARN", f"{date}: データ取得に失敗: {e}")
        return None

# 日次のヘルスケアデータ取得
def fetch_daily_data(garmin_client, date_obj):
    date_str = date_obj.strftime("%Y-%m-%d")
    stats = fetch_from_garmin(garmin_client.get_user_summary, date_str, date_str)
    if not stats:
        return None

    # Optional: Body composition data
    body_comp = fetch_from_garmin(garmin_client.get_body_composition, date_str, date_str) or {}

    # 必要なデータを抽出して加工
    data_dict = {
        "calendarDate": date_str,
        "steps": stats.get("totalSteps"),
        "distance_m": stats.get("totalDistanceMeters"),
        "sleep_hours": round(stats.get("sleepingSeconds", 0) / 3600, 2),
        "weight": round((body_comp.get("weight", stats.get("weight", 0)) / 1000), 2)
        if body_comp.get("weight", stats.get("weight", 0)) > 1000 else stats.get("weight", 0),
        # 他のデータも同様に...
    }
    return data_dict

# メイン処理
def main():
    # 環境変数チェック
    if not SA_KEY_VALUE:
        log_message("ERROR", "環境変数 'SA_KEY' が設定されていません。")
        sys.exit(1)

    # スプレッドシート接続
    spreadsheet = my_garmin_common.get_spreadsheet(SA_KEY_VALUE)
    if not spreadsheet:
        log_message("ERROR", "スプレッドシート接続に失敗。")
        sys.exit(1)

    # Garminログイン
    garmin = get_garmin_client()
    if not garmin:
        sys.exit(1)

    # データ取得
    # 環境変数 TARGET_DATE が指定されていれば、その日だけを処理する
    env_target_date = os.getenv("TARGET_DATE")
    log_message("DEBUG", f"受け取った日付指定(TARGET_DATE): '{env_target_date}'")
    
    if env_target_date:
        try:
            target_date = datetime.datetime.strptime(env_target_date, "%Y-%m-%d").date()
            target_dates = [target_date]
            log_message("INFO", f"🎯 指定日付モード: {target_date} のデータを取得します。")
        except ValueError:
            log_message("ERROR", f"❌ 日付形式エラー (YYYY-MM-DD): {env_target_date}")
            sys.exit(1)
    else:
        # 通常モード: 昨日 -> 今日
        today = datetime.datetime.now(ZoneInfo("Asia/Tokyo")).date()
        target_dates = [today - datetime.timedelta(days=i) for i in reversed(range(2))]

    new_data_list = []

    for date in target_dates:
        data = fetch_daily_data(garmin, date)
        if data:
            new_data_list.append(data)
        time.sleep(1)  # API制限防止のスリープ

    # データアップロード
    for data in new_data_list:
        my_garmin_common.update_daily_summary(spreadsheet, data)
        my_garmin_common.append_to_log(spreadsheet, data)
        log_message("INFO", f"{data['calendarDate']}: データアップロード成功。")

    # 最後にLogシートをソートして整える
    my_garmin_common.sort_log_sheet(spreadsheet)

if __name__ == "__main__":
    main()