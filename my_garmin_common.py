import pandas as pd
import os
import json
import datetime
from zoneinfo import ZoneInfo
from google.oauth2.service_account import Credentials
import gspread

# --- 定数定義 ---
SPREADSHEET_ID = '15CCDjcBCqSWYacPWf_RNXTBdJZ33x6pXAc1PhwPfkiY'

# スプレッドシートの列順序と日本語名の定義
COLUMN_MAP = {
    'timestamp': '実行日時',
    'calendarDate': '対象日付',
    # --- 活動量 ---
    'steps': '歩数',
    'distance_m': '移動距離',
    'floors_ascended': '上昇階数',
    'active_calories': '活動カロリー',
    'total_calories': '総カロリー',
    # --- 心拍・ストレス ---
    'heart_rate': '安静時心拍',
    'max_heart_rate': '最大心拍',
    'min_heart_rate': '最小心拍',
    'stress': 'ストレス',
    'body_battery': 'BodyBattery',
    # --- 睡眠 ---
    'sleep_hours': '睡眠時間',
    # --- 体組成 ---
    'weight': '体重',
    'bmi': 'BMI',
    'body_fat_pct': '体脂肪率',
    'muscle_pct': '筋肉率',
    'visceral_fat': '内臓脂肪',
    'metabolism': '基礎代謝',
    'bone_mass': '骨量',
    'water_ml': '水分摂取',
    'moderate_minutes': '中強度運動',
    'vigorous_minutes': '高強度運動',
}

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

def get_spreadsheet(sa_key_value):
    """スプレッドシートオブジェクトを取得して返す"""
    client = get_google_creds(sa_key_value)
    if not client:
        return None

    try:
        # サービスアカウントのメールアドレスを表示（権限確認用）
        try:
            print(f"ℹ️  Service Account Email: {client.auth.service_account_email}")
        except:
            pass

        return client.open_by_key(SPREADSHEET_ID)
    except Exception as e:
        print(f"❌ Spreadsheet Connection Error: {e}")
        if "403" in str(e):
            print("💡 ヒント: 上記のサービスアカウントのメールアドレスをコピーし、")
            print("   スプレッドシートの「共有」ボタンから「編集者」として追加してください。")
        return None

def update_daily_summary(spreadsheet, data_dict):
    """daily_summary シート（マスタ）を更新する"""
    SUMMARY_SHEET_NAME = 'daily_summary'
    try:
        try:
            sheet = spreadsheet.worksheet(SUMMARY_SHEET_NAME)
        except:
            print(f"Creating new sheet: {SUMMARY_SHEET_NAME}")
            sheet = spreadsheet.add_worksheet(title=SUMMARY_SHEET_NAME, rows=1000, cols=20)

        print(f"Updating {SUMMARY_SHEET_NAME}...")
        
        # 既存データを取得
        df_current = sheet_to_df(sheet)
        
        # 作業用データを作成（日付がない場合は今日を入れる）
        # JSTの現在日時を取得
        now_jst = datetime.datetime.now(ZoneInfo("Asia/Tokyo"))
        work_data = data_dict.copy()
        input_date_str = str(work_data.get('calendarDate') or "").strip()

        # 日付文字列の整形 (yyyy-mm-dd HH:mm:ss -> yyyy-mm-dd)
        target_date_str = ""
        if not input_date_str:
            target_date_str = now_jst.strftime("%Y-%m-%d")
        elif len(input_date_str) > 10:
            # 時刻が含まれている場合は日付部分だけ抽出
            target_date_str = input_date_str[:10]
        else:
            # 日付のみとみなす
            target_date_str = input_date_str

        work_data['calendarDate'] = target_date_str
        
        # 新しいデータをDataFrame化
        df_new = pd.DataFrame([work_data])
        df_new['calendarDate'] = df_new['calendarDate'].astype(str)
        df_new = df_new.set_index('calendarDate')
        
        # timestampはマスタには不要なので削除（もしあれば）
        if 'timestamp' in df_new.columns:
            df_new = df_new.drop(columns=['timestamp'])

        # マージとソート
        df_updated = df_new.combine_first(df_current)
        df_updated = df_updated.sort_index(ascending=False)
        
        # 保存
        df_to_sheet(sheet, df_updated)
        
    except Exception as e:
        print(f"⚠️ Daily Summary Update Failed: {e}")

def append_to_log(spreadsheet, data_dict):
    """Logシート（Sheet1）へデータを追記する"""
    print("Appending to Log Sheet...")
    try:
        sheet = spreadsheet.sheet1
        
        # 1. ヘッダー更新
        # COLUMN_MAPの順番通りにキーを並べる
        ordered_keys = list(COLUMN_MAP.keys())
        header_row = [f"{COLUMN_MAP[k]}({k})" for k in ordered_keys]
        
        try:
            sheet.update('A1', [header_row])
        except Exception as e:
            print(f"⚠️ Header Update Failed: {e}")

        # 2. データ行の作成
        # timestampを現在時刻に更新
        row_data = data_dict.copy()
        row_data['timestamp'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # calendarDateがない場合は今日の日付を入れる
        if not row_data.get('calendarDate'):
            row_data['calendarDate'] = datetime.date.today().strftime("%Y-%m-%d")
        
        # COLUMN_MAPの順番通りに値を並べる（不足キーはNone）
        values = [row_data.get(k) for k in ordered_keys]
        
        sheet.append_row(values)
        print("✅ Log sheet updated.")
        
    except Exception as e:
        print(f"❌ Log Sheet Append Error: {e}")
