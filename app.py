#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TAAAC 出退勤管理システム
"""

import os
import json
import uuid
import io
import base64
from datetime import datetime, timezone, timedelta
from pathlib import Path

from flask import Flask, request, jsonify, render_template_string, redirect, url_for, session

try:
    import qrcode
    import gspread
    from google.oauth2 import service_account
except ImportError as e:
    print(f"パッケージ不足: {e}")
    print("pip3 install flask 'qrcode[pil]' gspread google-auth")
    exit(1)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "taaac-timecard-secret-2026")
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(hours=24)

JST = timezone(timedelta(hours=9))
SPREADSHEET_ID = "1zLlshmq5AK1SoSG0Ezc1KhQbarI6g5_0H3lOwSYJOJU"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = Path(__file__).parent / "service_account.json"
STAFF_FILE = Path(__file__).parent / "staff.json"
ADMIN_PASSWORD = "taaac2026"

# ========== Google Sheets ==========

def get_sheets_client():
    # 環境変数からサービスアカウントJSON取得（Render用）
    sa_json = os.environ.get("SERVICE_ACCOUNT_JSON")
    if sa_json:
        import tempfile
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            f.write(sa_json)
            tmp_path = f.name
        creds = service_account.Credentials.from_service_account_file(tmp_path, scopes=SCOPES)
        os.unlink(tmp_path)
    elif SERVICE_ACCOUNT_FILE.exists():
        creds = service_account.Credentials.from_service_account_file(
            str(SERVICE_ACCOUNT_FILE), scopes=SCOPES
        )
    else:
        return None
    return gspread.authorize(creds)

# ========== スタッフデータ管理 ==========

STAFF_SHEET_NAME = "スタッフ"

def load_staff():
    gc = get_sheets_client()
    if gc:
        try:
            wb = gc.open_by_key(SPREADSHEET_ID)
            try:
                ws = wb.worksheet(STAFF_SHEET_NAME)
            except gspread.WorksheetNotFound:
                ws = wb.add_worksheet(title=STAFF_SHEET_NAME, rows=100, cols=4)
                ws.update([["スタッフID", "名前", "時給", "交通費"]], "A1")
                ws.format("A1:D1", {
                    "textFormat": {"bold": True},
                    "backgroundColor": {"red": 0.2, "green": 0.2, "blue": 0.2},
                    "horizontalAlignment": "CENTER"
                })
                # staff.jsonがあれば初回移行
                if STAFF_FILE.exists():
                    existing = json.loads(STAFF_FILE.read_text(encoding="utf-8"))
                    rows = [[sid, s["name"], s["wage"], s.get("transport", 0)] for sid, s in existing.items()]
                    if rows:
                        ws.append_rows(rows, value_input_option="USER_ENTERED")
                return json.loads(STAFF_FILE.read_text(encoding="utf-8")) if STAFF_FILE.exists() else {}
            rows = ws.get_all_values()
            staff = {}
            for row in rows[1:]:
                if len(row) >= 3 and row[0]:
                    staff[row[0]] = {
                        "name": row[1],
                        "wage": int(row[2]) if row[2] else 0,
                        "transport": int(row[3]) if len(row) > 3 and row[3] else 0,
                    }
            return staff
        except Exception as e:
            print(f"スタッフ読み込みエラー: {e}")
    # フォールバック: ローカルファイル
    if STAFF_FILE.exists():
        return json.loads(STAFF_FILE.read_text(encoding="utf-8"))
    return {}

def save_staff(data):
    gc = get_sheets_client()
    if gc:
        try:
            wb = gc.open_by_key(SPREADSHEET_ID)
            try:
                ws = wb.worksheet(STAFF_SHEET_NAME)
            except gspread.WorksheetNotFound:
                ws = wb.add_worksheet(title=STAFF_SHEET_NAME, rows=100, cols=4)
            ws.clear()
            ws.update([["スタッフID", "名前", "時給", "交通費"]], "A1")
            ws.format("A1:D1", {
                "textFormat": {"bold": True},
                "backgroundColor": {"red": 0.2, "green": 0.2, "blue": 0.2},
                "horizontalAlignment": "CENTER"
            })
            rows = [[sid, s["name"], s["wage"], s.get("transport", 0)] for sid, s in data.items()]
            if rows:
                ws.append_rows(rows, value_input_option="USER_ENTERED")
            return
        except Exception as e:
            print(f"スタッフ保存エラー: {e}")
    # フォールバック: ローカルファイル
    STAFF_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

WEEKDAYS_JP = ["月", "火", "水", "木", "金", "土", "日"]

# 列定義: A=日付, B=曜日, C=シフト開始, D=シフト終了, E=出勤, F=退勤, G=合計時間(h), H=給与, I=交通費, J=合計
# 行1: スタッフ名情報, 行2: 時給情報, 行3: ヘッダー, 行4以降: データ
HEADER_ROW = ["日付", "曜日", "シフト開始", "シフト終了", "出勤時間", "退勤時間", "合計時間(h)", "給与(円)", "交通費(円)", "合計(円)"]
DATA_START_ROW = 4  # 1-indexed

OWNER_EMAIL = "taaac.yokohama@gmail.com"
# 月別スプレッドシートIDキャッシュ: (year, month) -> spreadsheet_id
_monthly_ss_cache = {}

# サービスアカウントのDrive容量不足のため、手動作成済みスプレッドシートのIDを登録
_PRECREATED_SS = {
    (2026, 6): "1hWuPrBnueFL1xiAzLTq5ex8FUEq5msA0fPZkBy5owkI",
    (2026, 7): "1SubXg-EQ9wRREhc97-MAqJ7G8efKNdJT83AzmnBUE28",
}

def get_monthly_spreadsheet_title(year, month):
    return f"TAAAC出退勤_{year}-{month:02d}"

def get_or_create_monthly_spreadsheet(gc, year, month):
    """月別スプレッドシートを取得または新規作成してオーナーに共有"""
    cache_key = (year, month)

    # 手動作成済みIDが登録されていればそれを使う
    if cache_key not in _monthly_ss_cache and cache_key in _PRECREATED_SS:
        _monthly_ss_cache[cache_key] = _PRECREATED_SS[cache_key]

    if cache_key in _monthly_ss_cache:
        try:
            return gc.open_by_key(_monthly_ss_cache[cache_key])
        except Exception:
            pass

    title = get_monthly_spreadsheet_title(year, month)
    try:
        wb = gc.open(title)
        _monthly_ss_cache[cache_key] = wb.id
        return wb
    except gspread.SpreadsheetNotFound:
        pass

    # 新規作成
    wb = gc.create(title)
    _monthly_ss_cache[cache_key] = wb.id

    # オーナーに共有
    try:
        wb.share(OWNER_EMAIL, perm_type="user", role="writer", notify=False)
    except Exception as e:
        print(f"共有エラー: {e}")

    return wb

def get_or_create_staff_sheet(gc, staff_name, hourly_wage=None, year=None, month=None):
    now = datetime.now(JST)
    y = year or now.year
    m = month or now.month
    wb = get_or_create_monthly_spreadsheet(gc, y, m)

    try:
        ws = wb.worksheet(staff_name)
        if hourly_wage is not None:
            ws.update([[hourly_wage]], "B2")
    except gspread.WorksheetNotFound:
        # 最初のシートをリネームするか新規追加
        sheets = wb.worksheets()
        if len(sheets) == 1 and sheets[0].title in ("シート1", "Sheet1"):
            ws = sheets[0]
            ws.update_title(staff_name)
        else:
            ws = wb.add_worksheet(title=staff_name, rows=200, cols=10)
        ws.update([["スタッフ名", staff_name]], "A1")
        ws.update([["時給", hourly_wage or ""]], "A2")
        ws.update([HEADER_ROW], "A3")
        ws.format("A1:B2", {"textFormat": {"bold": True}})
        ws.format("A3:J3", {
            "textFormat": {"bold": True, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}},
            "backgroundColor": {"red": 0.2, "green": 0.2, "blue": 0.2},
            "horizontalAlignment": "CENTER"
        })
    return ws, wb

def find_next_row(ws):
    all_vals = ws.col_values(1)
    return len(all_vals) + 1

def add_monthly_summary_if_needed(ws, year, month):
    date_col = ws.col_values(1)
    month_prefix = f"{year:04d}-{month:02d}"
    rows_in_month = [i+1 for i, v in enumerate(date_col) if v.startswith(month_prefix)]
    if not rows_in_month:
        return
    last_row = max(rows_in_month)
    first_row = min(rows_in_month)
    next_after_last = last_row + 1
    vals_after = ws.row_values(next_after_last) if next_after_last <= len(date_col)+1 else []
    summary_label = f"{month}月 合計"
    sf_hours = f"=SUM(G{first_row}:G{last_row})"
    sf_pay   = f"=SUM(H{first_row}:H{last_row})"
    sf_trans = f"=SUM(I{first_row}:I{last_row})"
    sf_total = f"=SUM(J{first_row}:J{last_row})"

    row_data = [[summary_label, "", "", "", "", "", sf_hours, sf_pay, sf_trans, sf_total]]
    if vals_after and vals_after[0] == summary_label:
        ws.update(row_data, f"A{next_after_last}", value_input_option="USER_ENTERED")
    else:
        ws.append_row(row_data[0], value_input_option="USER_ENTERED")
        summary_row = find_next_row(ws) - 1
        ws.format(f"A{summary_row}:J{summary_row}", {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.85, "green": 0.93, "blue": 1.0}
        })

def record_clock_in_to_sheet(staff_name, clock_in_dt, hourly_wage, shift_start=None, shift_end=None):
    """出勤時点でスプシに記録（退勤欄は空）"""
    gc = get_sheets_client()
    if not gc:
        return None, "Google Sheets未認証"

    ws, wb = get_or_create_staff_sheet(gc, staff_name, hourly_wage, clock_in_dt.year, clock_in_dt.month)

    weekday = WEEKDAYS_JP[clock_in_dt.weekday()]
    date_str = clock_in_dt.strftime("%Y-%m-%d")
    shift_start_str = shift_start.strftime("%H:%M") if shift_start else ""
    shift_end_str = shift_end.strftime("%H:%M") if shift_end else ""

    row_data = [date_str, weekday, shift_start_str, shift_end_str,
                clock_in_dt.strftime("%H:%M"), "", "", "", "", ""]

    all_dates = ws.col_values(1)
    if all_dates and "月 合計" in all_dates[-1]:
        insert_pos = len(all_dates)
        ws.insert_row(row_data, insert_pos)
        return insert_pos, None
    else:
        ws.append_row(row_data)
        row_num = len(ws.col_values(1))
        return row_num, None

def record_clock_out_to_sheet(staff_name, clock_in_dt, clock_out_dt, hourly_wage, transport, row_num, shift_start=None, shift_end=None):
    """退勤時にスプシの出勤行を更新（計算列は数式で記入し手修正に対応）"""
    gc = get_sheets_client()
    if not gc:
        return False, "Google Sheets未認証"

    ws, wb = get_or_create_staff_sheet(gc, staff_name, hourly_wage, clock_in_dt.year, clock_in_dt.month)

    weekday = WEEKDAYS_JP[clock_in_dt.weekday()]
    date_str = clock_in_dt.strftime("%Y-%m-%d")
    shift_start_str = shift_start.strftime("%H:%M") if shift_start else ""
    shift_end_str = shift_end.strftime("%H:%M") if shift_end else ""
    clock_in_str = clock_in_dt.strftime("%H:%M")
    clock_out_str = clock_out_dt.strftime("%H:%M")

    # G列: シフト開始がある場合はMAX(出勤,シフト開始)〜退勤の時間、なければ出勤〜退勤
    # =IF(F{n}="","",IF(C{n}<>"",MAX(0,(F{n}-MAX(E{n},C{n}))*24),MAX(0,(F{n}-E{n})*24)))
    hours_formula = (
        f'=IF(F{row_num}="","",IF(C{row_num}<>"",'
        f'MAX(0,(TIMEVALUE(F{row_num})-MAX(TIMEVALUE(E{row_num}),TIMEVALUE(C{row_num})))*24),'
        f'MAX(0,(TIMEVALUE(F{row_num})-TIMEVALUE(E{row_num}))*24)))'
    )
    # H列: 給与 = 合計時間 × 時給(B2)
    pay_formula = f'=IF(G{row_num}="","",ROUND(G{row_num}*$B$2))'
    # J列: 合計 = 給与 + 交通費
    total_formula = f'=IF(H{row_num}="","",H{row_num}+IF(I{row_num}="",0,I{row_num}))'

    ws.update(
        [[date_str, weekday, shift_start_str, shift_end_str,
          clock_in_str, clock_out_str,
          hours_formula, pay_formula, transport, total_formula]],
        f"A{row_num}",
        value_input_option="USER_ENTERED"
    )
    add_monthly_summary_if_needed(ws, clock_in_dt.year, clock_in_dt.month)
    return True, None

def get_open_clockin(staff_name):
    """今日の未退勤行（退勤時間が空）を返す"""
    gc = get_sheets_client()
    if not gc:
        return None
    try:
        now = datetime.now(JST)
        wb = get_or_create_monthly_spreadsheet(gc, now.year, now.month)
        try:
            ws = wb.worksheet(staff_name)
        except gspread.WorksheetNotFound:
            return None
        today = now.strftime("%Y-%m-%d")
        all_rows = ws.get_all_values()
        for i, row in enumerate(all_rows[3:], start=4):
            if len(row) >= 5 and row[0] == today and row[4] and not row[5]:
                return {
                    "row_num": i,
                    "clock_in": row[4],
                    "shift_start": row[2],
                    "shift_end": row[3],
                }
    except Exception:
        pass
    return None

def get_monthly_records(staff_name, year, month):
    """スタッフの月別勤務データをスプシから取得"""
    gc = get_sheets_client()
    if not gc:
        return None, "Google Sheets未認証"
    try:
        wb = get_or_create_monthly_spreadsheet(gc, year, month)
        try:
            ws = wb.worksheet(staff_name)
        except gspread.WorksheetNotFound:
            # 旧形式（単一スプシ）にフォールバック
            try:
                old_wb = gc.open_by_key(SPREADSHEET_ID)
                ws = old_wb.worksheet(staff_name)
            except Exception:
                return [], None
    except Exception as e:
        return None, str(e)

    month_prefix = f"{year:04d}-{month:02d}"
    all_rows = ws.get_all_values()
    records = []
    for row in all_rows[3:]:
        if len(row) >= 1 and row[0].startswith(month_prefix):
            records.append(row)
    return records, None

# ========== セッション管理 ==========
active_sessions = {}  # staff_id -> {"clock_in": datetime, "shift_start": datetime|None, "shift_end": datetime|None}

# ========== HTML テンプレート ==========

PUNCH_HTML = """
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>打刻 - {{ name }}</title>
<link rel="apple-touch-icon" href="/static/logo.png">
<link rel="icon" href="/static/favicon.ico">
<meta name="apple-mobile-web-app-title" content="TAAAC打刻">
<meta name="apple-mobile-web-app-capable" content="yes">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, sans-serif; background: #f5f5f5; min-height: 100vh; display: flex; align-items: center; justify-content: center; }
  .card { background: white; border-radius: 20px; padding: 40px 32px; max-width: 360px; width: 90%; text-align: center; box-shadow: 0 4px 24px rgba(0,0,0,0.1); }
  .avatar { width: 80px; height: 80px; border-radius: 50%; background: #e8f4fd; display: flex; align-items: center; justify-content: center; font-size: 32px; margin: 0 auto 16px; }
  h1 { font-size: 22px; color: #222; margin-bottom: 4px; }
  .time { font-size: 36px; font-weight: bold; color: #333; margin: 20px 0; }
  .date { color: #888; font-size: 14px; margin-bottom: 24px; }
  .status { font-size: 14px; padding: 6px 16px; border-radius: 20px; display: inline-block; margin-bottom: 24px; }
  .status.in { background: #e8f7ee; color: #27ae60; }
  .status.out { background: #fef9e7; color: #f39c12; }
  .btn { width: 100%; padding: 18px; border: none; border-radius: 14px; font-size: 18px; font-weight: bold; cursor: pointer; transition: 0.15s; }
  .btn-in { background: #27ae60; color: white; }
  .btn-out { background: #e74c3c; color: white; }
  .btn:active { transform: scale(0.97); }
  .msg { margin-top: 20px; padding: 14px; border-radius: 10px; font-size: 15px; }
  .msg.success { background: #e8f7ee; color: #27ae60; }
  .msg.error { background: #fdecea; color: #e74c3c; }
  .mypage-link { display: block; margin-top: 18px; color: #3498db; font-size: 14px; text-decoration: none; }
  .shift-form { margin-bottom: 14px; text-align: left; }
  .shift-form label { font-size: 12px; color: #888; display: block; margin-bottom: 6px; }
  .shift-row { display: flex; gap: 8px; align-items: center; }
  .shift-row span { font-size: 13px; color: #666; }
  .shift-row input[type=time] { flex: 1; padding: 10px 8px; border: 1px solid #ddd; border-radius: 8px; font-size: 15px; text-align: center; }
  .btn-row { display: flex; gap: 10px; margin-top: 4px; }
  .btn-row .btn { flex: 1; width: auto; padding: 16px 8px; font-size: 16px; }
  .status.in { background: #e8f7ee; color: #27ae60; }
  .status.out { background: #fef9e7; color: #f39c12; }
</style>
</head>
<body>
<div class="card">
  <div class="avatar">👤</div>
  <h1>{{ name }} さん</h1>
  <div class="time" id="clock">--:--:--</div>
  <div class="date" id="date"></div>
  {% if open_clockin %}
  <div class="status in">● 出勤中（{{ open_clockin.clock_in }}〜）</div>
  {% else %}
  <div class="status out">○ 未出勤</div>
  {% endif %}
  {% if message %}
  <div class="msg {{ msg_type }}">{{ message }}</div>
  {% else %}
  <form method="post">
    <div class="shift-form">
      <label>シフト時間（出勤時に入力・任意）</label>
      <div class="shift-row">
        <input type="time" name="shift_start">
        <span>〜</span>
        <input type="time" name="shift_end">
      </div>
    </div>
    <div class="btn-row">
      <button class="btn btn-in" type="submit" name="action" value="in">出勤</button>
      <button class="btn btn-out" type="submit" name="action" value="out">退勤</button>
    </div>
  </form>
  {% endif %}
  <a class="mypage-link" href="/mypage/{{ staff_id }}">📊 今月の実績を見る</a>
</div>
<script>
function updateClock() {
  const now = new Date();
  const pad = n => String(n).padStart(2, '0');
  document.getElementById('clock').textContent =
    pad(now.getHours()) + ':' + pad(now.getMinutes()) + ':' + pad(now.getSeconds());
  const days = ['日','月','火','水','木','金','土'];
  document.getElementById('date').textContent =
    now.getFullYear() + '年' + (now.getMonth()+1) + '月' + now.getDate() + '日（' + days[now.getDay()] + '）';
}
setInterval(updateClock, 1000);
updateClock();
{% if message %}
setTimeout(() => { window.location.href = window.location.pathname; }, 3000);
{% endif %}
</script>
</body>
</html>
"""

MYPAGE_HTML = """
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{{ name }}さんの実績</title>
<link rel="apple-touch-icon" href="/static/logo.png">
<link rel="icon" href="/static/favicon.ico">
<meta name="apple-mobile-web-app-capable" content="yes">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, sans-serif; background: #f5f5f5; min-height: 100vh; }
  header { background: #1a1a2e; color: white; padding: 16px 20px; display: flex; justify-content: space-between; align-items: center; }
  header h1 { font-size: 17px; }
  .back { color: #aaa; font-size: 13px; text-decoration: none; }
  .container { max-width: 600px; margin: 0 auto; padding: 20px 16px; }
  .summary-cards { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; margin-bottom: 20px; }
  .summary-card { background: white; border-radius: 14px; padding: 18px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  .summary-card .label { font-size: 12px; color: #888; margin-bottom: 6px; }
  .summary-card .value { font-size: 22px; font-weight: bold; color: #222; }
  .summary-card.highlight { background: #1a1a2e; }
  .summary-card.highlight .label { color: #aaa; }
  .summary-card.highlight .value { color: white; }
  .section { background: white; border-radius: 14px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  h2 { font-size: 15px; color: #333; margin-bottom: 14px; border-bottom: 2px solid #f0f0f0; padding-bottom: 8px; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th { background: #1a1a2e; color: white; padding: 8px 6px; text-align: center; font-weight: normal; }
  td { padding: 9px 6px; text-align: center; border-bottom: 1px solid #f5f5f5; }
  tr:last-child td { border-bottom: none; }
  .sat { color: #3498db; }
  .sun { color: #e74c3c; }
  .month-nav { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; }
  .month-nav a { color: #3498db; text-decoration: none; font-size: 14px; padding: 6px 14px; border: 1px solid #3498db; border-radius: 8px; }
  .month-title { font-size: 16px; font-weight: bold; color: #333; }
  .error { background: #fdecea; color: #e74c3c; padding: 14px; border-radius: 10px; text-align: center; font-size: 14px; }
  .empty { color: #aaa; text-align: center; padding: 30px; font-size: 14px; }
</style>
</head>
<body>
<header>
  <h1>{{ name }}さんの実績</h1>
  <a href="/punch/{{ staff_id }}" class="back">← 打刻に戻る</a>
</header>
<div class="container">

  <div class="month-nav">
    <a href="/mypage/{{ staff_id }}?year={{ prev_year }}&month={{ prev_month }}">← 前月</a>
    <span class="month-title">{{ year }}年{{ month }}月</span>
    <a href="/mypage/{{ staff_id }}?year={{ next_year }}&month={{ next_month }}">次月 →</a>
  </div>

  {% if error %}
  <div class="error">{{ error }}</div>
  {% elif not records %}
  <div class="empty">この月の記録はありません</div>
  {% else %}

  <div class="summary-cards">
    <div class="summary-card">
      <div class="label">出勤日数</div>
      <div class="value">{{ work_days }}日</div>
    </div>
    <div class="summary-card">
      <div class="label">合計時間</div>
      <div class="value">{{ "%.1f"|format(total_hours) }}h</div>
    </div>
    <div class="summary-card">
      <div class="label">給与合計</div>
      <div class="value">¥{{ "{:,}".format(total_pay) }}</div>
    </div>
    <div class="summary-card highlight">
      <div class="label">交通費込み合計</div>
      <div class="value">¥{{ "{:,}".format(total_all) }}</div>
    </div>
  </div>

  <div class="section">
    <h2>勤務記録</h2>
    <table>
      <tr>
        <th>日付</th><th>曜日</th><th>シフト</th><th>出勤</th><th>退勤</th><th>時間</th><th>給与</th>{% if has_transport %}<th>交通費</th>{% endif %}
      </tr>
      {% for r in records %}
      <tr>
        <td>{{ r.date[5:] }}</td>
        <td class="{{ 'sat' if r.weekday == '土' else ('sun' if r.weekday == '日' else '') }}">{{ r.weekday }}</td>
        <td style="font-size:12px;color:#888;">{% if r.shift_start %}{{ r.shift_start }}〜{{ r.shift_end }}{% else %}—{% endif %}</td>
        <td>{{ r.clock_in }}</td>
        <td>{{ r.clock_out }}</td>
        <td>{{ r.hours }}h</td>
        <td>¥{{ "{:,}".format(r.pay) }}</td>
        {% if has_transport %}<td>¥{{ "{:,}".format(r.transport) }}</td>{% endif %}
      </tr>
      {% endfor %}
    </table>
  </div>
  {% endif %}

</div>
</body>
</html>
"""

ADMIN_LOGIN_HTML = """
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>管理者ログイン</title>
<link rel="icon" href="/static/favicon.ico">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, sans-serif; background: #1a1a2e; min-height: 100vh; display: flex; align-items: center; justify-content: center; }
  .card { background: white; border-radius: 16px; padding: 40px 32px; max-width: 360px; width: 90%; text-align: center; }
  h1 { font-size: 20px; margin-bottom: 24px; color: #333; }
  input { width: 100%; padding: 14px; border: 2px solid #ddd; border-radius: 10px; font-size: 16px; margin-bottom: 16px; outline: none; }
  input:focus { border-color: #3498db; }
  .btn { width: 100%; padding: 14px; background: #3498db; color: white; border: none; border-radius: 10px; font-size: 16px; font-weight: bold; cursor: pointer; }
  .error { color: #e74c3c; font-size: 14px; margin-bottom: 12px; }
</style>
</head>
<body>
<div class="card">
  <h1>管理者ログイン</h1>
  {% if error %}<div class="error">{{ error }}</div>{% endif %}
  <form method="post">
    <input type="password" name="password" placeholder="パスワード" autofocus>
    <button class="btn" type="submit">ログイン</button>
  </form>
</div>
</body>
</html>
"""

ADMIN_HTML = """
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>出退勤管理 - TAAAC</title>
<link rel="icon" href="/static/favicon.ico">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, sans-serif; background: #f5f5f5; }
  header { background: #1a1a2e; color: white; padding: 16px 24px; display: flex; justify-content: space-between; align-items: center; }
  header h1 { font-size: 18px; }
  .logout { color: #aaa; font-size: 13px; text-decoration: none; }
  .container { max-width: 900px; margin: 0 auto; padding: 24px 16px; }
  .section { background: white; border-radius: 14px; padding: 24px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  h2 { font-size: 16px; color: #333; margin-bottom: 18px; border-bottom: 2px solid #f0f0f0; padding-bottom: 10px; }
  .staff-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 12px; }
  .staff-card { border: 2px solid #f0f0f0; border-radius: 12px; padding: 16px; position: relative; }
  .staff-card.active { border-color: #27ae60; background: #f0faf4; }
  .staff-name { font-weight: bold; font-size: 15px; margin-bottom: 4px; }
  .staff-wage { color: #888; font-size: 13px; }
  .staff-transport { color: #8e44ad; font-size: 12px; margin-top: 2px; }
  .staff-status { font-size: 12px; margin-top: 8px; }
  .status-in { color: #27ae60; }
  .status-out { color: #aaa; }
  .qr-btn { display: block; margin-top: 8px; padding: 6px 12px; background: #3498db; color: white; border: none; border-radius: 6px; font-size: 12px; cursor: pointer; text-decoration: none; text-align: center; }
  .edit-btn { display: block; margin-top: 6px; padding: 6px 12px; background: #f39c12; color: white; border: none; border-radius: 6px; font-size: 12px; cursor: pointer; text-align: center; width: 100%; }
  .add-form { display: flex; gap: 10px; flex-wrap: wrap; }
  .add-form input, .add-form select { flex: 1; min-width: 100px; padding: 10px 14px; border: 2px solid #ddd; border-radius: 8px; font-size: 14px; outline: none; }
  .add-form input:focus { border-color: #3498db; }
  .add-form button { padding: 10px 20px; background: #27ae60; color: white; border: none; border-radius: 8px; font-size: 14px; font-weight: bold; cursor: pointer; white-space: nowrap; }
  .delete-btn { position: absolute; top: 8px; right: 8px; background: none; border: none; color: #ddd; cursor: pointer; font-size: 16px; }
  .delete-btn:hover { color: #e74c3c; }
  .modal { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.7); z-index: 100; align-items: center; justify-content: center; }
  .modal.open { display: flex; }
  .modal-inner { background: white; border-radius: 16px; padding: 32px; text-align: center; max-width: 340px; width: 90%; }
  .modal-inner img { max-width: 240px; margin: 16px auto; display: block; }
  .modal-inner h3 { margin-bottom: 8px; }
  .modal-inner p { color: #888; font-size: 13px; margin-bottom: 16px; word-break: break-all; }
  .close-btn { padding: 10px 24px; background: #333; color: white; border: none; border-radius: 8px; cursor: pointer; }
  .edit-modal-inner { background: white; border-radius: 16px; padding: 28px; max-width: 340px; width: 90%; }
  .edit-modal-inner h3 { margin-bottom: 18px; font-size: 16px; }
  .edit-modal-inner input { width: 100%; padding: 10px 14px; border: 2px solid #ddd; border-radius: 8px; font-size: 14px; outline: none; margin-bottom: 12px; }
  .edit-modal-inner input:focus { border-color: #3498db; }
  .edit-modal-inner label { display: block; font-size: 12px; color: #888; margin-bottom: 4px; }
  .edit-save-btn { width: 100%; padding: 12px; background: #27ae60; color: white; border: none; border-radius: 8px; font-size: 15px; font-weight: bold; cursor: pointer; margin-top: 4px; }
  .edit-cancel-btn { width: 100%; padding: 10px; background: #f5f5f5; color: #888; border: none; border-radius: 8px; font-size: 14px; cursor: pointer; margin-top: 8px; }
  .alert { padding: 12px 16px; border-radius: 8px; margin-bottom: 16px; font-size: 14px; }
  .alert-success { background: #e8f7ee; color: #27ae60; }
  .alert-error { background: #fdecea; color: #e74c3c; }
</style>
</head>
<body>
<header>
  <h1>TAAAC 出退勤管理</h1>
  <div style="display:flex;gap:8px;">
    <a href="/summary" class="logout">📊 集計</a>
    <a href="/admin/update_delivery_list" class="logout">🚚 配達リスト更新</a>
    <a href="https://drive.google.com/drive/u/0/folders/1-OSqYzfAcA8XE31GqoT7V1c-JsZgtldc" class="logout" target="_blank">Drive</a>
    <a href="/admin/logout" class="logout">ログアウト</a>
  </div>
</header>
<div class="container">

  {% if flash_msg %}
  <div class="alert alert-{{ flash_type }}">{{ flash_msg }}</div>
  {% endif %}

  <!-- スタッフ一覧 -->
  <div class="section">
    <h2>スタッフ一覧</h2>
    {% if not staff %}
    <p style="color:#aaa;font-size:14px;">スタッフがまだ登録されていません</p>
    {% else %}
    <div class="staff-grid">
      {% for sid, s in staff.items() %}
      <div class="staff-card {{ 'active' if sid in active else '' }}">
        <button class="delete-btn" onclick="deleteStaff('{{ sid }}', '{{ s.name }}')">✕</button>
        <div class="staff-name">{{ s.name }}</div>
        <div class="staff-wage">時給 ¥{{ s.wage | int | format_number }}</div>
        {% if s.transport %}
        <div class="staff-transport">交通費 ¥{{ s.transport | int | format_number }}/日</div>
        {% endif %}
        <div class="staff-status">
          {% if sid in active %}
          <span class="status-in">● 出勤中（{{ active[sid] }}〜）</span>
          {% else %}
          <span class="status-out">○ 未出勤</span>
          {% endif %}
        </div>
        <a class="qr-btn" href="#" onclick="showQR('{{ sid }}', '{{ s.name }}')">QRコード表示</a>
        <button class="edit-btn" onclick="showEdit('{{ sid }}', '{{ s.name }}', {{ s.wage }}, {{ s.transport or 0 }})">編集</button>
      </div>
      {% endfor %}
    </div>
    {% endif %}
  </div>

  <!-- 手動打刻追加 -->
  <div class="section">
    <h2>手動で打刻を追加</h2>
    <form method="post" action="/admin/manual_punch">
      <div class="add-form">
        <select name="staff_id" required style="flex:2;padding:10px 14px;border:2px solid #ddd;border-radius:8px;font-size:14px;outline:none;">
          <option value="">スタッフを選択</option>
          {% for sid, s in staff.items() %}
          <option value="{{ sid }}">{{ s.name }}</option>
          {% endfor %}
        </select>
        <div style="flex:1.5;display:flex;flex-direction:column;gap:3px;">
          <label style="font-size:11px;color:#888;">日付</label>
          <input type="date" name="date" required style="width:100%;">
        </div>
        <div style="flex:1;display:flex;flex-direction:column;gap:3px;">
          <label style="font-size:11px;color:#888;">シフト開始</label>
          <input type="time" name="shift_start" style="width:100%;">
        </div>
        <div style="flex:1;display:flex;flex-direction:column;gap:3px;">
          <label style="font-size:11px;color:#888;">シフト終了</label>
          <input type="time" name="shift_end" style="width:100%;">
        </div>
        <div style="flex:1;display:flex;flex-direction:column;gap:3px;">
          <label style="font-size:11px;color:#888;">出勤打刻</label>
          <input type="time" name="clock_in" required style="width:100%;">
        </div>
        <div style="flex:1;display:flex;flex-direction:column;gap:3px;">
          <label style="font-size:11px;color:#888;">退勤打刻</label>
          <input type="time" name="clock_out" required style="width:100%;">
        </div>
        <button type="submit" style="background:#8e44ad;">追加</button>
      </div>
    </form>
  </div>

  <!-- 外部URL設定 -->
  <div class="section">
    <h2>外部URL設定（ngrok用）</h2>
    <form method="post" action="/admin/set_base_url">
      <div class="add-form">
        <input type="text" name="base_url" placeholder="https://xxxx.ngrok-free.app" value="{{ base_url }}" style="flex:3">
        <button type="submit">保存</button>
      </div>
    </form>
    <p style="font-size:12px;color:#aaa;margin-top:8px;">QRコードのURLに使われます。ngrokのURLを貼り付けてください。</p>
  </div>

  <!-- スタッフ追加 -->
  <div class="section">
    <h2>スタッフ追加</h2>
    <form method="post" action="/admin/add_staff">
      <div class="add-form">
        <input type="text" name="name" placeholder="名前" required>
        <input type="number" name="wage" placeholder="時給（円）" min="900" required>
        <input type="number" name="transport" placeholder="交通費/日（円）" min="0" value="0">
        <button type="submit">追加</button>
      </div>
    </form>
  </div>

</div>

<!-- QRモーダル -->
<div class="modal" id="qrModal">
  <div class="modal-inner">
    <h3 id="qrName"></h3>
    <p id="qrUrl"></p>
    <img id="qrImg" src="" alt="QR">
    <br>
    <button class="close-btn" onclick="closeQR()">閉じる</button>
  </div>
</div>

<!-- 編集モーダル -->
<div class="modal" id="editModal">
  <div class="edit-modal-inner">
    <h3>✏️ スタッフ編集</h3>
    <form method="post" action="/admin/edit_staff">
      <input type="hidden" name="staff_id" id="editStaffId">
      <label>名前</label>
      <input type="text" id="editName" name="name" required>
      <label>時給（円）</label>
      <input type="number" id="editWage" name="wage" min="900" required>
      <label>交通費/日（円）</label>
      <input type="number" id="editTransport" name="transport" min="0" value="0">
      <button type="submit" class="edit-save-btn">保存</button>
    </form>
    <button class="edit-cancel-btn" onclick="closeEdit()">キャンセル</button>
  </div>
</div>

<script>
const BASE_URL = "{{ base_url }}";
function showQR(sid, name) {
  document.getElementById('qrName').textContent = name + ' さん';
  const url = BASE_URL + '/punch/' + sid + '?skip=1';
  document.getElementById('qrUrl').textContent = url;
  document.getElementById('qrImg').src = '/admin/qr/' + sid + '?base_url=' + encodeURIComponent(BASE_URL);
  document.getElementById('qrModal').classList.add('open');
}
function closeQR() { document.getElementById('qrModal').classList.remove('open'); }
function showEdit(sid, name, wage, transport) {
  document.getElementById('editStaffId').value = sid;
  document.getElementById('editName').value = name;
  document.getElementById('editWage').value = wage;
  document.getElementById('editTransport').value = transport;
  document.getElementById('editModal').classList.add('open');
}
function closeEdit() { document.getElementById('editModal').classList.remove('open'); }
function deleteStaff(sid, name) {
  if (confirm(name + ' さんを削除しますか？')) {
    fetch('/admin/delete_staff/' + sid, {method: 'POST'}).then(() => location.reload());
  }
}
</script>
</body>
</html>
"""

# ========== Jinja フィルター ==========

@app.template_filter('format_number')
def format_number(value):
    return f"{int(value):,}"

# ========== ルート ==========

@app.after_request
def add_ngrok_header(response):
    response.headers["ngrok-skip-browser-warning"] = "true"
    return response

def _session_key(staff_id, field):
    return f"punch_{staff_id}_{field}"

def _restore_from_cookie(staff_id):
    """サーバー再起動後もクッキーから打刻状態を復元する"""
    if staff_id in active_sessions:
        return
    clock_in_iso = session.get(_session_key(staff_id, "clock_in"))
    if clock_in_iso:
        try:
            clock_in_dt = datetime.fromisoformat(clock_in_iso)
            shift_start_iso = session.get(_session_key(staff_id, "shift_start"))
            shift_end_iso = session.get(_session_key(staff_id, "shift_end"))
            shift_start_dt = datetime.fromisoformat(shift_start_iso) if shift_start_iso else None
            shift_end_dt = datetime.fromisoformat(shift_end_iso) if shift_end_iso else None
            row_num = session.get(_session_key(staff_id, "row_num"))
            active_sessions[staff_id] = {
                "clock_in": clock_in_dt,
                "shift_start": shift_start_dt,
                "shift_end": shift_end_dt,
                "row_num": int(row_num) if row_num else None,
            }
        except Exception:
            pass

def _save_to_cookie(staff_id, clock_in_dt, shift_start_dt, shift_end_dt, row_num=None):
    session[_session_key(staff_id, "clock_in")] = clock_in_dt.isoformat()
    session[_session_key(staff_id, "shift_start")] = shift_start_dt.isoformat() if shift_start_dt else ""
    session[_session_key(staff_id, "shift_end")] = shift_end_dt.isoformat() if shift_end_dt else ""
    session[_session_key(staff_id, "row_num")] = str(row_num) if row_num else ""

def _clear_cookie(staff_id):
    for field in ("clock_in", "shift_start", "shift_end", "row_num"):
        session.pop(_session_key(staff_id, field), None)

@app.route("/punch/<staff_id>", methods=["GET", "POST"])
def punch(staff_id):
    staff = load_staff()
    if staff_id not in staff:
        return "スタッフが見つかりません", 404

    s = staff[staff_id]
    now = datetime.now(JST)
    message = None
    msg_type = None

    if request.method == "POST":
        action = request.form.get("action")
        shift_start_str = request.form.get("shift_start", "").strip()
        shift_end_str = request.form.get("shift_end", "").strip()
        today = now.date()
        shift_start_dt = None
        shift_end_dt = None
        if shift_start_str:
            try:
                shift_start_dt = datetime.strptime(f"{today} {shift_start_str}", "%Y-%m-%d %H:%M").replace(tzinfo=JST)
            except ValueError:
                pass
        if shift_end_str:
            try:
                shift_end_dt = datetime.strptime(f"{today} {shift_end_str}", "%Y-%m-%d %H:%M").replace(tzinfo=JST)
            except ValueError:
                pass

        if action == "in":
            row_num, err = record_clock_in_to_sheet(s["name"], now, s["wage"], shift_start_dt, shift_end_dt)
            if row_num:
                message = f"出勤しました ✓ {now.strftime('%H:%M')}"
            else:
                message = f"出勤しました ✓ {now.strftime('%H:%M')}（スプシ未接続: {err}）"
            msg_type = "success"

        elif action == "out":
            open_row = get_open_clockin(s["name"])
            if open_row:
                clock_in_dt = datetime.strptime(
                    f"{now.strftime('%Y-%m-%d')} {open_row['clock_in']}", "%Y-%m-%d %H:%M"
                ).replace(tzinfo=JST)
                shift_start_dt = datetime.strptime(
                    f"{now.strftime('%Y-%m-%d')} {open_row['shift_start']}", "%Y-%m-%d %H:%M"
                ).replace(tzinfo=JST) if open_row["shift_start"] else None
                shift_end_dt = datetime.strptime(
                    f"{now.strftime('%Y-%m-%d')} {open_row['shift_end']}", "%Y-%m-%d %H:%M"
                ).replace(tzinfo=JST) if open_row["shift_end"] else None
                transport = s.get("transport", 0)
                ok, err = record_clock_out_to_sheet(
                    s["name"], clock_in_dt, now, s["wage"], transport,
                    open_row["row_num"], shift_start_dt, shift_end_dt
                )
                if ok:
                    pay_start = shift_start_dt if (shift_start_dt and clock_in_dt < shift_start_dt) else clock_in_dt
                    hours = (now - pay_start).total_seconds() / 3600
                    pay = round(hours * s["wage"])
                    total = pay + transport
                    message = f"退勤しました ✓ {now.strftime('%H:%M')}｜{hours:.1f}h｜¥{total:,}"
                else:
                    message = f"退勤記録 ✓ {now.strftime('%H:%M')}（スプシ未接続: {err}）"
            else:
                message = "本日の出勤記録が見つかりません"
                msg_type = "error"
            if msg_type != "error":
                msg_type = "success"

    open_clockin = get_open_clockin(s["name"]) if not message else None

    return render_template_string(
        PUNCH_HTML,
        staff_id=staff_id,
        name=s["name"],
        open_clockin=open_clockin,
        message=message,
        msg_type=msg_type
    )

@app.route("/mypage/<staff_id>")
def mypage(staff_id):
    staff = load_staff()
    if staff_id not in staff:
        return "スタッフが見つかりません", 404

    s = staff[staff_id]
    now = datetime.now(JST)

    year = int(request.args.get("year", now.year))
    month = int(request.args.get("month", now.month))

    # 前月・次月計算
    if month == 1:
        prev_year, prev_month = year - 1, 12
    else:
        prev_year, prev_month = year, month - 1
    if month == 12:
        next_year, next_month = year + 1, 1
    else:
        next_year, next_month = year, month + 1

    records_raw, error = get_monthly_records(s["name"], year, month)

    records = []
    total_hours = 0
    total_pay = 0
    total_transport = 0

    if records_raw:
        for row in records_raw:
            # row: [日付, 曜日, シフト開始, シフト終了, 出勤, 退勤, 合計h, 給与, 交通費, 合計]
            try:
                hours = float(row[6]) if len(row) > 6 and row[6] else 0
                pay = int(row[7]) if len(row) > 7 and row[7] else 0
                transport = int(row[8]) if len(row) > 8 and row[8] else 0
                records.append({
                    "date": row[0],
                    "weekday": row[1] if len(row) > 1 else "",
                    "shift_start": row[2] if len(row) > 2 else "",
                    "shift_end": row[3] if len(row) > 3 else "",
                    "clock_in": row[4] if len(row) > 4 else "",
                    "clock_out": row[5] if len(row) > 5 else "",
                    "hours": hours,
                    "pay": pay,
                    "transport": transport,
                })
                total_hours += hours
                total_pay += pay
                total_transport += transport
            except Exception:
                pass

    has_transport = any(r["transport"] > 0 for r in records)
    total_all = total_pay + total_transport

    return render_template_string(
        MYPAGE_HTML,
        staff_id=staff_id,
        name=s["name"],
        year=year,
        month=month,
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        records=records,
        work_days=len(records),
        total_hours=total_hours,
        total_pay=total_pay,
        total_transport=total_transport,
        total_all=total_all,
        has_transport=has_transport,
        error=error
    )

@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    error = None
    if request.method == "POST":
        if request.form.get("password") == ADMIN_PASSWORD:
            session["admin"] = True
            return redirect(url_for("admin"))
        error = "パスワードが違います"
    return render_template_string(ADMIN_LOGIN_HTML, error=error)

@app.route("/admin/logout")
def admin_logout():
    session.pop("admin", None)
    return redirect(url_for("admin_login"))

@app.route("/admin")
def admin():
    if not session.get("admin"):
        return redirect(url_for("admin_login"))

    staff = load_staff()
    active = {sid: v["clock_in"].strftime("%H:%M") for sid, v in active_sessions.items()}
    host = request.host
    base_url = session.get("base_url") or os.environ.get("BASE_URL") or f"https://{host}"
    flash_msg = session.pop("flash_msg", None)
    flash_type = session.pop("flash_type", "success")

    return render_template_string(
        ADMIN_HTML,
        staff=staff,
        active=active,
        base_url=base_url,
        flash_msg=flash_msg,
        flash_type=flash_type
    )

@app.route("/admin/add_staff", methods=["POST"])
def add_staff():
    if not session.get("admin"):
        return redirect(url_for("admin_login"))

    name = request.form.get("name", "").strip()
    wage = request.form.get("wage", "0")
    transport = request.form.get("transport", "0")

    if not name:
        session["flash_msg"] = "名前を入力してください"
        session["flash_type"] = "error"
        return redirect(url_for("admin"))

    staff = load_staff()
    staff_id = str(uuid.uuid4())[:8]
    staff[staff_id] = {"name": name, "wage": int(wage), "transport": int(transport)}
    save_staff(staff)

    gc = get_sheets_client()
    if gc:
        get_or_create_staff_sheet(gc, name)

    session["flash_msg"] = f"{name} さんを追加しました（時給¥{int(wage):,} / 交通費¥{int(transport):,}/日）"
    session["flash_type"] = "success"
    return redirect(url_for("admin"))

@app.route("/admin/edit_staff", methods=["POST"])
def edit_staff():
    if not session.get("admin"):
        return redirect(url_for("admin_login"))

    staff_id = request.form.get("staff_id")
    name = request.form.get("name", "").strip()
    wage = int(request.form.get("wage", "0"))
    transport = int(request.form.get("transport", "0"))

    staff = load_staff()
    if staff_id not in staff:
        session["flash_msg"] = "スタッフが見つかりません"
        session["flash_type"] = "error"
        return redirect(url_for("admin"))

    staff[staff_id]["name"] = name
    staff[staff_id]["wage"] = wage
    staff[staff_id]["transport"] = transport
    save_staff(staff)

    session["flash_msg"] = f"{name} さんの情報を更新しました"
    session["flash_type"] = "success"
    return redirect(url_for("admin"))

@app.route("/admin/delete_staff/<staff_id>", methods=["POST"])
def delete_staff(staff_id):
    if not session.get("admin"):
        return jsonify({"error": "unauthorized"}), 401

    staff = load_staff()
    if staff_id in staff:
        del staff[staff_id]
        save_staff(staff)
        active_sessions.pop(staff_id, None)
    return jsonify({"ok": True})

@app.route("/admin/manual_punch", methods=["POST"])
def manual_punch():
    if not session.get("admin"):
        return redirect(url_for("admin_login"))

    staff = load_staff()
    staff_id = request.form.get("staff_id")
    date_str = request.form.get("date")
    clock_in_str = request.form.get("clock_in")
    clock_out_str = request.form.get("clock_out")
    shift_start_str = request.form.get("shift_start", "").strip()
    shift_end_str = request.form.get("shift_end", "").strip()

    if not all([staff_id, date_str, clock_in_str, clock_out_str]) or staff_id not in staff:
        session["flash_msg"] = "入力内容を確認してください"
        session["flash_type"] = "error"
        return redirect(url_for("admin"))

    s = staff[staff_id]
    clock_in_dt = datetime.strptime(f"{date_str} {clock_in_str}", "%Y-%m-%d %H:%M").replace(tzinfo=JST)
    clock_out_dt = datetime.strptime(f"{date_str} {clock_out_str}", "%Y-%m-%d %H:%M").replace(tzinfo=JST)
    shift_start_dt = datetime.strptime(f"{date_str} {shift_start_str}", "%Y-%m-%d %H:%M").replace(tzinfo=JST) if shift_start_str else None
    shift_end_dt = datetime.strptime(f"{date_str} {shift_end_str}", "%Y-%m-%d %H:%M").replace(tzinfo=JST) if shift_end_str else None

    if clock_out_dt <= clock_in_dt:
        session["flash_msg"] = "退勤時刻は出勤時刻より後にしてください"
        session["flash_type"] = "error"
        return redirect(url_for("admin"))

    transport = s.get("transport", 0)
    row_num, err = record_clock_in_to_sheet(s["name"], clock_in_dt, s["wage"], shift_start_dt, shift_end_dt)
    if row_num:
        ok, err = record_clock_out_to_sheet(s["name"], clock_in_dt, clock_out_dt, s["wage"], transport, row_num, shift_start_dt, shift_end_dt)
    else:
        ok = False
    pay_start = shift_start_dt if (shift_start_dt and clock_in_dt < shift_start_dt) else clock_in_dt
    hours = (clock_out_dt - pay_start).total_seconds() / 3600
    pay = round(hours * s["wage"])
    total = pay + transport

    if ok:
        session["flash_msg"] = f"{s['name']} {date_str} {clock_in_str}〜{clock_out_str}（{hours:.1f}h / ¥{total:,}）を記録しました"
        session["flash_type"] = "success"
    else:
        session["flash_msg"] = f"スプシへの記録に失敗しました: {err}"
        session["flash_type"] = "error"

    return redirect(url_for("admin"))

@app.route("/admin/set_base_url", methods=["POST"])
def set_base_url():
    if not session.get("admin"):
        return redirect(url_for("admin_login"))
    base_url = request.form.get("base_url", "").strip().rstrip("/")
    session["base_url"] = base_url
    session["flash_msg"] = "外部URLを設定しました"
    session["flash_type"] = "success"
    return redirect(url_for("admin"))

@app.route("/admin/qr/<staff_id>")
def qr_image(staff_id):
    if not session.get("admin"):
        return redirect(url_for("admin_login"))

    base_url = request.args.get("base_url") or session.get("base_url") or os.environ.get("BASE_URL") or f"https://{request.host}"
    url = f"{base_url}/punch/{staff_id}?skip=1"

    img = qrcode.make(url)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)

    from flask import send_file
    return send_file(buf, mimetype="image/png")

@app.route("/summary")
@app.route("/summary/<int:year>/<int:month>")
def summary(year=None, month=None):
    if not session.get("admin"):
        return redirect(url_for("admin_login"))
    now = datetime.now(JST)
    year = year or now.year
    month = month or now.month

    staff = load_staff()
    gc = get_sheets_client()

    rows_by_staff = {}
    error_msg = None
    if gc:
        try:
            wb = get_or_create_monthly_spreadsheet(gc, year, month)
            for s in staff.values():
                name = s["name"]
                try:
                    ws = wb.worksheet(name)
                    all_rows = ws.get_all_values()
                    # 日付フォーマットが YYYY-MM-DD / YYYY/MM/DD どちらでも対応
                    month_patterns = (
                        f"{year:04d}-{month:02d}",
                        f"{year:04d}/{month:02d}",
                        f"{year}/{month}/",
                    )
                    data = [r for r in all_rows[3:] if r and r[0] and
                            any(r[0].startswith(p) for p in month_patterns)
                            and "合計" not in r[0]]
                    rows_by_staff[name] = data
                except Exception:
                    rows_by_staff[name] = []
        except Exception as e:
            error_msg = str(e)

    # 集計
    summary_data = []
    total_hours = 0.0
    total_pay = 0
    total_transport = 0
    for s in staff.values():
        name = s["name"]
        rows = rows_by_staff.get(name, [])
        count = len(rows)
        h = sum(float(r[6]) for r in rows if len(r) > 6 and r[6] and r[6] not in ("", "-"))
        pay = sum(int(r[7]) for r in rows if len(r) > 7 and r[7] and r[7].lstrip("-").isdigit())
        tr = sum(int(r[8]) for r in rows if len(r) > 8 and r[8] and r[8].lstrip("-").isdigit())
        total_hours += h
        total_pay += pay
        total_transport += tr
        if count > 0:
            summary_data.append({"name": name, "count": count, "hours": h, "pay": pay, "transport": tr})

    # 前月・次月
    prev_month = month - 1 if month > 1 else 12
    prev_year = year if month > 1 else year - 1
    next_month = month + 1 if month < 12 else 1
    next_year = year if month < 12 else year + 1

    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>月次集計 {year}年{month}月</title>
<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: -apple-system, sans-serif; background: #f0f2f5; min-height: 100vh; }}
header {{ background: #1a1a2e; color: white; padding: 16px 20px; display: flex; align-items: center; justify-content: space-between; }}
header h1 {{ font-size: 17px; font-weight: 700; }}
.back-btn {{ color: #aaa; text-decoration: none; font-size: 14px; }}
.nav {{ display: flex; align-items: center; justify-content: center; gap: 16px; padding: 16px; background: white; border-bottom: 1px solid #eee; }}
.nav a {{ color: #3498db; text-decoration: none; font-size: 20px; padding: 4px 12px; }}
.nav-title {{ font-size: 17px; font-weight: bold; color: #222; }}
.container {{ padding: 16px; max-width: 480px; margin: 0 auto; }}
.total-card {{ background: linear-gradient(135deg, #1a1a2e, #16213e); color: white; border-radius: 16px; padding: 24px; margin-bottom: 16px; }}
.total-card h2 {{ font-size: 13px; color: #aaa; margin-bottom: 16px; letter-spacing: 1px; }}
.total-grid {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; }}
.total-item {{ text-align: center; }}
.total-label {{ font-size: 11px; color: #aaa; margin-bottom: 4px; }}
.total-value {{ font-size: 20px; font-weight: bold; }}
.total-value.green {{ color: #2ecc71; }}
.total-value.blue {{ color: #3498db; }}
.total-value.yellow {{ color: #f1c40f; }}
.section-title {{ font-size: 13px; font-weight: bold; color: #888; letter-spacing: 1px; margin-bottom: 10px; }}
.member-card {{ background: white; border-radius: 12px; padding: 16px; margin-bottom: 10px; display: flex; align-items: center; gap: 14px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); }}
.member-avatar {{ width: 44px; height: 44px; border-radius: 50%; background: #e8f4fd; display: flex; align-items: center; justify-content: center; font-size: 18px; flex-shrink: 0; }}
.member-info {{ flex: 1; min-width: 0; }}
.member-name {{ font-size: 15px; font-weight: bold; color: #222; margin-bottom: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
.member-stats {{ font-size: 12px; color: #888; }}
.member-right {{ text-align: right; flex-shrink: 0; }}
.member-pay {{ font-size: 16px; font-weight: bold; color: #27ae60; }}
.member-hours {{ font-size: 12px; color: #aaa; }}
.badge {{ display: inline-block; background: #3498db; color: white; border-radius: 20px; padding: 2px 10px; font-size: 12px; font-weight: bold; }}
.empty {{ text-align: center; padding: 40px; color: #aaa; font-size: 14px; }}
</style>
</head>
<body>
<header>
  <h1>📊 月次集計</h1>
  <a href="/admin" class="back-btn">管理画面</a>
</header>
<div class="nav">
  <a href="/summary/{prev_year}/{prev_month}">←</a>
  <span class="nav-title">{year}年{month}月</span>
  <a href="/summary/{next_year}/{next_month}">→</a>
</div>
<div class="container">
"""
    if error_msg:
        html += f'<div style="background:#fdecea;color:#e74c3c;padding:12px;border-radius:8px;margin-bottom:16px;font-size:13px;">{error_msg}</div>'

    html += f"""
  <div class="total-card">
    <h2>TODAY'S MONTH TOTAL</h2>
    <div class="total-grid">
      <div class="total-item">
        <div class="total-label">合計給与</div>
        <div class="total-value green">¥{total_pay:,}</div>
      </div>
      <div class="total-item">
        <div class="total-label">合計時間</div>
        <div class="total-value blue">{total_hours:.1f}h</div>
      </div>
      <div class="total-item">
        <div class="total-label">交通費</div>
        <div class="total-value yellow">¥{total_transport:,}</div>
      </div>
    </div>
  </div>
  <div class="section-title">メンバー別</div>
"""
    if not summary_data:
        html += '<div class="empty">出勤データがありません</div>'
    else:
        for d in sorted(summary_data, key=lambda x: x["pay"], reverse=True):
            html += f"""
  <div class="member-card">
    <div class="member-avatar">👤</div>
    <div class="member-info">
      <div class="member-name">{d["name"]}</div>
      <div class="member-stats"><span class="badge">{d["count"]}回</span> &nbsp; {d["hours"]:.1f}h &nbsp; 交通費 ¥{d["transport"]:,}</div>
    </div>
    <div class="member-right">
      <div class="member-pay">¥{d["pay"]:,}</div>
      <div class="member-hours">給与</div>
    </div>
  </div>"""

    html += "\n</div></body></html>"
    return html

DELIVERY_LIST_SS_ID = "1I5EnWoK3h7pnpdhcvsSHsmzAe9d2NbmUYOKuwrnhRtI"

@app.route("/admin/update_delivery_list")
def update_delivery_list():
    if not session.get("admin"):
        return redirect(url_for("admin_login"))

    gc = get_sheets_client()
    if not gc:
        return "Google Sheets未認証", 500

    try:
        delivery_wb = gc.open_by_key(DELIVERY_LIST_SS_ID)
    except Exception as e:
        return f"配達リストスプレッドシートを開けません: {e}", 500

    staff = load_staff()
    results = []

    # 対象月: _PRECREATED_SS に登録済みの月 + 今月
    now = datetime.now(JST)
    target_months = set(_PRECREATED_SS.keys())
    target_months.add((now.year, now.month))

    for (year, month) in sorted(target_months):
        sheet_name = f"{year}-{month:02d}"
        month_prefix = f"{year:04d}-{month:02d}"

        # 全スタッフのその月のデータを収集
        all_records = []
        try:
            src_wb = get_or_create_monthly_spreadsheet(gc, year, month)
        except Exception as e:
            results.append(f"{sheet_name}: スプレッドシート取得失敗 ({e})")
            continue

        for s in staff.values():
            name = s["name"]
            try:
                ws = src_wb.worksheet(name)
                rows = ws.get_all_values()
                for row in rows[3:]:
                    if not row or not row[0]:
                        continue
                    if "合計" in str(row[0]):
                        continue
                    date_val = row[0]
                    if not (date_val.startswith(month_prefix) or
                            date_val.startswith(f"{year}/{month:02d}") or
                            date_val.startswith(f"{year}/{month}/")):
                        continue
                    # [日付, 曜日, シフト開始, シフト終了, 出勤, 退勤, 合計h, 給与, 交通費, 合計]
                    clock_in  = row[4] if len(row) > 4 else ""
                    clock_out = row[5] if len(row) > 5 else ""
                    pay       = row[7] if len(row) > 7 else ""
                    if not clock_in:
                        continue
                    # 合計時間を小数点第2位切り上げ
                    import math
                    raw_hours = row[6] if len(row) > 6 else ""
                    try:
                        hours = f"{math.ceil(float(raw_hours) * 100) / 100:.2f}"
                    except (ValueError, TypeError):
                        hours = raw_hours
                    all_records.append([date_val, row[1] if len(row) > 1 else "", name, clock_in, clock_out, hours, pay])
            except Exception:
                continue

        if not all_records:
            results.append(f"{sheet_name}: データなし")
            continue

        # 日付順にソート
        all_records.sort(key=lambda r: r[0])

        # シートを取得または作成
        try:
            dst_ws = delivery_wb.worksheet(sheet_name)
            dst_ws.clear()
        except gspread.exceptions.WorksheetNotFound:
            dst_ws = delivery_wb.add_worksheet(title=sheet_name, rows=500, cols=7)

        header = [["日付", "曜日", "スタッフ名", "出勤", "退勤", "合計時間(h)", "給与(円)"]]
        dst_ws.update(header + all_records, "A1", value_input_option="USER_ENTERED")

        # ヘッダー書式
        dst_ws.format("A1:G1", {
            "textFormat": {"bold": True, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}},
            "backgroundColor": {"red": 0.2, "green": 0.2, "blue": 0.2},
            "horizontalAlignment": "CENTER"
        })

        import time; time.sleep(1)
        results.append(f"{sheet_name}: {len(all_records)}件を書き込み")

    # デフォルトシート(シート1)があれば削除
    try:
        default = delivery_wb.worksheet("シート1")
        delivery_wb.del_worksheet(default)
    except Exception:
        pass

    link = f'<a href="https://docs.google.com/spreadsheets/d/{DELIVERY_LIST_SS_ID}" target="_blank">配達スタッフ出勤リストを開く</a>'
    return "<br>".join(results) + f"<br><br>{link}<br><br><a href='/admin'>管理画面に戻る</a>"

@app.route("/")
def index():
    return redirect(url_for("admin"))

@app.route("/admin/migrate_sheets")
def migrate_sheets():
    if not session.get("admin"):
        return redirect(url_for("admin_login"))

    gc = get_sheets_client()
    if not gc:
        return "Google Sheets未認証", 500

    staff = load_staff()
    wb = gc.open_by_key(SPREADSHEET_ID)
    results = []

    for ws in wb.worksheets():
        name = ws.title
        # スタッフシートと月次サマリー以外のシートを対象に
        if name in (STAFF_SHEET_NAME,):
            continue
        # スタッフ名と一致するシートのみ処理
        staff_info = next((s for s in staff.values() if s["name"] == name), None)
        if not staff_info:
            continue

        all_rows = ws.get_all_values()
        if not all_rows:
            results.append(f"{name}: 空のシート、スキップ")
            continue

        # フォーマット判定
        header = all_rows[0]
        if header and header[0] == "スタッフ名":
            results.append(f"{name}: 既に新フォーマット、スキップ")
            continue

        # 旧フォーマットのデータ抽出
        data_rows = []
        summary_labels = []
        for row in all_rows[1:]:
            if not row or not row[0]:
                continue
            # 月合計行の判定（旧: A=名前 C=月合計, または A=月合計）
            if len(row) > 2 and "月 合計" in str(row[2]):
                summary_labels.append(row[2])
                data_rows.append(("summary", row[2]))
                continue
            if "月 合計" in str(row[0]):
                summary_labels.append(row[0])
                data_rows.append(("summary", row[0]))
                continue

            # 旧12列フォーマット: [名前,時給,日付,曜日,シフト開始,シフト終了,出勤,退勤,合計h,給与,交通費,合計]
            if len(row) >= 12 and row[2].startswith("20"):
                data_rows.append(("data", [row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11]]))
            # 旧10列フォーマット: [名前,時給,日付,曜日,出勤,退勤,合計h,給与,交通費,合計]
            elif len(row) >= 10 and row[2].startswith("20"):
                data_rows.append(("data", [row[2], row[3], "", "", row[4], row[5], row[6], row[7], row[8], row[9]]))
            else:
                continue

        # シートを新フォーマットで書き直す
        ws.clear()
        ws.update([["スタッフ名", name]], "A1")
        ws.update([["時給", staff_info["wage"]]], "A2")
        ws.update([HEADER_ROW], "A3")
        ws.format("A1:B2", {"textFormat": {"bold": True}})
        ws.format("A3:J3", {
            "textFormat": {"bold": True, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}},
            "backgroundColor": {"red": 0.2, "green": 0.2, "blue": 0.2},
            "horizontalAlignment": "CENTER"
        })

        current_row = 4
        summary_row_nums = []
        for item in data_rows:
            if item[0] == "summary":
                ws.append_row([item[1]] + [""] * 9, value_input_option="USER_ENTERED")
                summary_row_nums.append(current_row)
            else:
                ws.append_row(item[1], value_input_option="USER_ENTERED")
            current_row += 1

        for r in summary_row_nums:
            ws.format(f"A{r}:J{r}", {
                "textFormat": {"bold": True},
                "backgroundColor": {"red": 0.85, "green": 0.93, "blue": 1.0}
            })

        data_count = len([d for d in data_rows if d[0] == "data"])
        results.append(f"{name}: {data_count}件移行完了")

    return "<br>".join(results) + "<br><br><a href='/admin'>管理画面に戻る</a>"

@app.route("/admin/migrate_june")
def migrate_june():
    if not session.get("admin"):
        return redirect(url_for("admin_login"))

    try:
        return _migrate_june_impl()
    except Exception as e:
        import traceback
        return f"エラー: {e}<br><pre>{traceback.format_exc()}</pre>", 500

def _migrate_june_impl():
    gc = get_sheets_client()
    if not gc:
        return "Google Sheets未認証", 500

    staff = load_staff()
    old_wb = gc.open_by_key(SPREADSHEET_ID)
    results = []

    # 2026-06用スプレッドシート（手動作成済み）
    JUNE_SS_ID = "1hWuPrBnueFL1xiAzLTq5ex8FUEq5msA0fPZkBy5owkI"
    new_wb = gc.open_by_key(JUNE_SS_ID)
    # キャッシュに登録してget_or_create_staff_sheetでも使えるようにする
    _monthly_ss_cache[(2026, 6)] = JUNE_SS_ID
    results.append(f"移行先スプレッドシート: TAAAC出退勤_2026-06 (ID: {JUNE_SS_ID})")

    for ws in old_wb.worksheets():
        name = ws.title
        if name in (STAFF_SHEET_NAME,):
            continue
        staff_info = next((s for s in staff.values() if s["name"] == name), None)
        if not staff_info:
            continue

        all_rows = ws.get_all_values()
        if not all_rows:
            continue

        # データ行の抽出（6月のみ）
        june_rows = []
        # 新フォーマット (row1=スタッフ名, row2=時給, row3=ヘッダー, row4+=データ)
        header = all_rows[0]
        if header and header[0] == "スタッフ名":
            data_start = 3  # 0-indexed: row index 3 = row 4
            for row in all_rows[data_start:]:
                if not row or not row[0]:
                    continue
                if "月 合計" in str(row[0]):
                    continue
                if row[0].startswith("2026-06"):
                    june_rows.append(("data", row[:10]))
        else:
            # 旧フォーマット
            for row in all_rows[1:]:
                if not row or not row[0]:
                    continue
                if "月 合計" in str(row[2] if len(row) > 2 else ""):
                    continue
                if len(row) >= 12 and row[2].startswith("2026-06"):
                    june_rows.append(("data", [row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11]]))
                elif len(row) >= 10 and row[2].startswith("2026-06"):
                    june_rows.append(("data", [row[2], row[3], "", "", row[4], row[5], row[6], row[7], row[8], row[9]]))

        if not june_rows:
            results.append(f"{name}: 6月データなし、スキップ")
            continue

        # 移行先に新しいシートを作成
        try:
            new_ws = new_wb.worksheet(name)
            new_ws.clear()
        except gspread.exceptions.WorksheetNotFound:
            new_ws = new_wb.add_worksheet(title=name, rows=200, cols=10)

        wage = staff_info["wage"]
        # ヘッダー + データを一括書き込み
        all_data = [
            ["スタッフ名", name],
            ["時給", wage],
            HEADER_ROW,
        ] + [item[1] + [""] * (10 - len(item[1])) for item in june_rows]
        new_ws.update(all_data, "A1", value_input_option="USER_ENTERED")
        import time; time.sleep(2)  # レート制限対策

        results.append(f"{name}: {len(june_rows)}件の6月データを移行完了")

    # デフォルトシート(Sheet1)があれば削除
    try:
        default_ws = new_wb.worksheet("Sheet1")
        new_wb.del_worksheet(default_ws)
        results.append("デフォルトSheet1を削除")
    except:
        pass

    results.append(f'<a href="https://docs.google.com/spreadsheets/d/{new_wb.id}" target="_blank">移行先スプレッドシートを開く</a>')
    return "<br>".join(results) + "<br><br><a href='/admin'>管理画面に戻る</a>"

if __name__ == "__main__":
    print("=" * 50)
    print("TAAAC 出退勤管理システム起動")
    print(f"管理画面: http://localhost:5001/admin")
    print(f"パスワード: {ADMIN_PASSWORD}")
    print("=" * 50)
    app.run(host="0.0.0.0", port=5001, debug=False)
