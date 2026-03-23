import pandas as pd
import os
import json
import datetime
from zoneinfo import ZoneInfo
from google.oauth2.service_account import Credentials
import gspread
from gspread.utils import rowcol_to_a1

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
    """daily_summary シート（マスタ）を更新する。テーブル形式を破壊しないように処理する。"""
    SUMMARY_SHEET_NAME = 'daily_summary'
    # 1-based index of the column to sort by (A=1, B=2, etc.)
    SORT_COLUMN_INDEX = 2 # calendarDate
    try:
        try:
            sheet = spreadsheet.worksheet(SUMMARY_SHEET_NAME)
        except:
            print(f"Creating new sheet: {SUMMARY_SHEET_NAME}")
            sheet = spreadsheet.add_worksheet(title=SUMMARY_SHEET_NAME, rows=1000, cols=20)

        print(f"Updating {SUMMARY_SHEET_NAME} (table-safe)...")

        # --- 1. 入力データの日付を正規化 ---
        # JSTの現在日時を取得
        now_jst = datetime.datetime.now(ZoneInfo("Asia/Tokyo"))
        work_data = data_dict.copy()
        input_date_str = str(work_data.get('calendarDate') or "").strip()

        if not input_date_str:
            target_date_str = now_jst.strftime("%Y-%m-%d")
        elif len(input_date_str) > 10:
            target_date_str = input_date_str[:10]
        else:
            target_date_str = input_date_str

        work_data['calendarDate'] = target_date_str
        if 'timestamp' in work_data:
            del work_data['timestamp']

        # --- 2. シートのヘッダーとデータを取得 ---
        all_data = sheet.get_all_values()
        if not all_data: # シートが完全に空の場合
            print(f"Sheet '{SUMMARY_SHEET_NAME}' is empty. Initializing...")
            # ヘッダーの順序をCOLUMN_MAPから（timestamp以外で）作成
            headers = [v for k, v in COLUMN_MAP.items() if k != 'timestamp']
            sheet.append_row(headers)
            # ヘッダーに対応するデータ行を作成
            new_row = [work_data.get(key) for key in COLUMN_MAP if key != 'timestamp']
            sheet.append_row(new_row)
            print(f"✨ Sheet '{sheet.title}' initialized and data inserted.")
            return

        headers = all_data[0]
        try:
            date_col_index = headers.index('対象日付') # 日本語ヘッダー名で検索
        except ValueError:
            print(f"❌ Critical Error: '対象日付' column not found in {SUMMARY_SHEET_NAME}.")
            return

        # --- 3. 更新対象の行を探索 ---
        target_row_number = -1
        for i, row in enumerate(all_data[1:]): # ヘッダーを除いて探索
            if len(row) > date_col_index and row[date_col_index] == target_date_str:
                target_row_number = i + 2 # 1-based index, and +1 for header
                break

        # --- 4. 更新または追記 ---
        if target_row_number > -1:
            # 既存行を更新
            print(f"Updating existing row {target_row_number} for date {target_date_str}...")
            current_row_values = sheet.row_values(target_row_number)
            # 更新後の行データを作成
            updated_values = []
            for i, header_name in enumerate(headers):
                # 逆引きしてキー名を取得
                key_name = next((k for k, v in COLUMN_MAP.items() if v == header_name), None)
                # 新しいデータに値があればそれを使う、なければ既存の値を使う
                if key_name and key_name in work_data and work_data[key_name] is not None:
                    updated_values.append(work_data[key_name])
                else:
                    updated_values.append(current_row_values[i] if i < len(current_row_values) else "")
            sheet.update(f'A{target_row_number}', [updated_values], value_input_option='USER_ENTERED')
        else:
            # 新しい行を追記
            print(f"Appending new row for date {target_date_str}...")
            new_row = [work_data.get(next((k for k, v in COLUMN_MAP.items() if v == h), None), "") for h in headers]
            sheet.append_row(new_row, value_input_option='USER_ENTERED')
            sheet.sort((date_col_index + 1, 'des')) # 追記後に日付で降順ソート
        
        print(f"✨ Sheet '{sheet.title}' updated.")
    except Exception as e:
        print(f"⚠️ Daily Summary Update Failed: {e}")

def append_to_log(spreadsheet, data_dict):
    """Logシート（Sheet1）へデータを追記する。テーブルが拡張されるようにヘッダーの直後に追加する。"""
    print("Appending to Log Sheet...")
    try:
        sheet = spreadsheet.sheet1
        
        # 1. ヘッダー更新
        # COLUMN_MAPの順番通りにキーを並べる
        ordered_keys = list(COLUMN_MAP.keys())
        header_row = [f"{COLUMN_MAP[k]}({k})" for k in ordered_keys]
        
        try:
            # A1セルが空の場合のみヘッダーを書き込む
            if not sheet.acell('A1').value:
                sheet.update('A1', [header_row])
        except Exception as e:
            print(f"⚠️ Header check/update failed: {e}")

        # 2. データ行の作成
        # timestampを現在時刻(JST)に更新
        row_data = data_dict.copy()
        now_jst = datetime.datetime.now(ZoneInfo("Asia/Tokyo"))
        row_data['timestamp'] = now_jst.strftime("%Y-%m-%d %H:%M:%S")
        
        # calendarDateがない場合は今日の日付を入れる
        if not row_data.get('calendarDate'):
            row_data['calendarDate'] = now_jst.strftime("%Y-%m-%d")
        
        # COLUMN_MAPの順番通りに値を並べる（不足キーはNone）
        values = [row_data.get(k) for k in ordered_keys]
        
        # ヘッダーの直後(2行目)に新しいログを挿入する
        # これにより、テーブルが確実に拡張され、最新のログが一番上に来る
        sheet.insert_row(values, 2, value_input_option='USER_ENTERED')
        print("✅ Log sheet updated (newest on top).")
        
    except Exception as e:
        print(f"❌ Log Sheet Append Error: {e}")

def sort_log_sheet(spreadsheet):
    """Logシートを実行日時(A列)降順 -> 対象日付(B列)降順でソートする。ヘッダー(1行目)は除外。"""
    print("Sorting Log Sheet...")
    try:
        sheet = spreadsheet.sheet1
        # データのある全範囲を取得
        all_values = sheet.get_all_values()
        num_rows = len(all_values)
        if num_rows < 2:
            print("ℹ️ No data to sort.")
            return

        num_cols = len(all_values[0])
        # ソート範囲を決定 (A2 から 右下のセルまで)
        sort_range = f"A2:{rowcol_to_a1(num_rows, num_cols)}"
        
        # 1列目(timestamp)を降順、2列目(calendarDate)を降順でソート
        sheet.sort((1, 'des'), (2, 'des'), range=sort_range)
        print("✅ Log sheet sorted (Timestamp DESC -> Date DESC).")
        
    except Exception as e:
        print(f"⚠️ Log Sheet Sort Error: {e}")
