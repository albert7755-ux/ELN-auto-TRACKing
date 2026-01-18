import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import re
from dateutil.relativedelta import relativedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- 設定網頁 ---
st.set_page_config(page_title="ELN 智能戰情室 (自動抓價版)", layout="wide")

# ==========================================
# 🔐 雲端機密讀取 (Gmail + LINE)
# ==========================================
try:
    # 嘗試讀取 LINE 設定
    LINE_ACCESS_TOKEN = st.secrets.get("LINE_ACCESS_TOKEN", "")
    MY_LINE_USER_ID = st.secrets.get("MY_LINE_USER_ID", "")
    
    # 嘗試讀取 Gmail 設定
    GMAIL_ACCOUNT = st.secrets.get("GMAIL_ACCOUNT", "")
    GMAIL_PASSWORD = st.secrets.get("GMAIL_PASSWORD", "")
    ADMIN_EMAIL = st.secrets.get("ADMIN_EMAIL", GMAIL_ACCOUNT)
except Exception:
    st.error("⚠️ Secrets 設定讀取異常，部分功能可能無法使用。")
    LINE_ACCESS_TOKEN = ""
    MY_LINE_USER_ID = ""
    GMAIL_ACCOUNT = ""
    GMAIL_PASSWORD = ""
    ADMIN_EMAIL = ""

# ==========================================
# 🔄 狀態初始化
# ==========================================
if 'last_processed_file' not in st.session_state:
    st.session_state['last_processed_file'] = None
if 'is_sent' not in st.session_state:
    st.session_state['is_sent'] = False

# --- 側邊欄 ---
with st.sidebar:
    st.header("⚙️ 設定中心")
    
    status_text = ""
    if LINE_ACCESS_TOKEN: status_text += "✅ LINE 連線 OK\n"
    if GMAIL_ACCOUNT: status_text += "✅ Email 連線 OK"
    if not status_text: status_text = "⚠️ 未設定連線金鑰"
    st.success(status_text)

    st.markdown("---")
    real_today = datetime.now()
    st.info(f"📅 今天日期：{real_today.strftime('%Y-%m-%d')}")
    st.caption("鎖定為真實日期")

    st.markdown("---")
    st.header("🔔 通知過濾")
    st.caption("程式會自動回溯抓取交易日價格，無須手動輸入。")
    
    lookback_days = st.slider("只通知幾天內發生的事件？", min_value=1, max_value=30, value=3)
    notify_ki_daily = st.checkbox("KI (跌破) 是否每天提醒？", value=True)

    st.info("💡 **小技巧**\n支援新版格式：自動將 `TSLA UW` 轉為 `TSLA` 並抓取交易日價格。")

# --- 函數區 ---

# 🌟 關鍵新增：代號清洗器 (把 Bloomberg 格式轉成 Yahoo 格式)
def clean_ticker_symbol(ticker):
    if pd.isna(ticker): return ""
    t = str(ticker).strip().upper()
    
    # 美股：去除 UW, UN, UQ, UP
    for suffix in [" UW", " UN", " UQ", " UP"]:
        if t.endswith(suffix): return t.replace(suffix, "")
    
    # 日股：JT -> .T
    if t.endswith(" JT"): return t.replace(" JT", ".T")
    
    # 台股：TT -> .TW (假設)
    if t.endswith(" TT"): return t.replace(" TT", ".TW")
    
    # 港股：HK -> .HK
    if t.endswith(" HK"): return t.replace(" HK", ".HK")
    
    return t

def send_line_push(target_user_id, message_text):
    if not LINE_ACCESS_TOKEN or not target_user_id: return False
    from linebot import LineBotApi
    from linebot.models import TextSendMessage
    try:
        uid = str(target_user_id).strip()
        if not uid.startswith("U") or len(uid) < 10: return False
        line_bot_api = LineBotApi(LINE_ACCESS_TOKEN)
        line_bot_api.push_message(uid, TextSendMessage(text=message_text))
        return True
    except Exception as e:
        print(f"LINE 發送失敗: {e}"); return False

def send_email_gmail(to_email, subject, body_text):
    if not GMAIL_ACCOUNT or not GMAIL_PASSWORD or not to_email: return False
    if "@" not in str(to_email): return False
    try:
        msg = MIMEMultipart()
        msg['From'] = GMAIL_ACCOUNT
        msg['To'] = str(to_email).strip()
        msg['Subject'] = subject
        msg.attach(MIMEText(body_text, 'plain'))
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(GMAIL_ACCOUNT, GMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"Email 發送失敗: {e}"); return False

def parse_nc_months(ko_type_str):
    if pd.isna(ko_type_str) or str(ko_type_str).strip() == "": return 1 
    match = re.search(r'NC(\d+)', str(ko_type_str), re.IGNORECASE)
    if match: return int(match.group(1))
    return 1 

def clean_percentage(val):
    if pd.isna(val) or str(val).strip() == "": return None
    try:
        s = str(val).replace('%', '').replace(',', '').strip()
        return float(s)
    except: return None

def clean_name_str(val):
    if pd.isna(val): return "貴賓"
    s = str(val).strip()
    if s.lower() == 'nan' or s == "": return "貴賓"
    return s

def find_col_index(columns, include_keywords, exclude_keywords=None):
    for idx, col_name in enumerate(columns):
        col_str = str(col_name).strip().lower()
        if exclude_keywords:
            if any(ex in col_str for ex in exclude_keywords): continue
        if any(inc in col_str for inc in include_keywords):
            return idx, col_name
    return None, None

# --- 主畫面 ---
st.title("📊 ELN 智能戰情室 - 自動抓價版")

uploaded_file = st.file_uploader("請上傳 Excel (支援新版無價格格式)", type=['xlsx', 'csv'], key="uploader")

if uploaded_file:
    if st.session_state['last_processed_file'] != uploaded_file.name:
        st.session_state['last_processed_file'] = uploaded_file.name
        st.session_state['is_sent'] = False

if uploaded_file is not None:
    try:
        # 1. 讀取檔案
        try:
            df = pd.read_excel(uploaded_file, sheet_name=0, header=0, engine='openpyxl')
        except:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file)

        df = df.dropna(how='all')
        # 簡單過濾標題行
        if df.iloc[0].astype(str).str.contains("進場價").any():
            df = df.iloc[1:].reset_index(drop=True)
            
        cols = df.columns.tolist()
        
        # 2. 欄位定位
        id_idx, _ = find_col_index(cols, ["債券", "代號", "id"]) or (0, "")
        strike_idx, _ = find_col_index(cols, ["strike", "執行", "履約"])
        ko_idx, _ = find_col_index(cols, ["ko", "提前"], exclude_keywords=["strike", "執行", "ki", "type"])
        ko_type_idx, _ = find_col_index(cols, ["ko類型", "ko type"]) or find_col_index(cols, ["類型", "type"], exclude_keywords=["ki", "ko"])
        ki_idx, _ = find_col_index(cols, ["ki", "下檔"], exclude_keywords=["ko", "type"])
        ki_type_idx, _ = find_col_index(cols, ["ki類型", "ki type"])
        t1_idx, _ = find_col_index(cols, ["標的1", "ticker 1"])
        
        trade_date_idx, _ = find_col_index(cols, ["交易日"])
        issue_date_idx, _ = find_col_index(cols, ["發行日"])
        final_date_idx, _ = find_col_index(cols, ["最終", "評價"])
        maturity_date_idx, _ = find_col_index(cols, ["到期", "maturity"])
        
        name_idx, _ = find_col_index(cols, ["理專", "姓名", "客戶"])
        line_id_idx, _ = find_col_index(cols, ["line_id", "lineid", "line user id", "uid"])
        email_idx, _ = find_col_index(cols, ["email", "e-mail", "mail", "信箱"])

        if t1_idx is None:
            st.error("❌ 無法辨識「標的1」欄位，請檢查 Excel 表頭。")
            st.stop()

        # 3. 建立標準化資料表
        clean_df = pd.DataFrame()
        clean_df['ID'] = df.iloc[:, id_idx]
        if name_idx is not None: clean_df['Name'] = df.iloc[:, name_idx].apply(clean_name_str)
        else: clean_df['Name'] = "貴賓"
        
        if line_id_idx is not None: clean_df['Line_ID'] = df.iloc[:, line_id_idx].astype(str).replace('nan', '').str.strip()
        else: clean_df['Line_ID'] = ""
        
        if email_idx is not None: clean_df['Email'] = df.iloc[:, email_idx].astype(str).replace('nan', '').str.strip()
        else: clean_df['Email'] = ""

        clean_df['TradeDate'] = pd.to_datetime(df.iloc[:, trade_date_idx], errors='coerce') if trade_date_idx else pd.NaT
        clean_df['IssueDate'] = pd.to_datetime(df.iloc[:, issue_date_idx], errors='coerce') if issue_date_idx else pd.Timestamp.min
        clean_df['ValuationDate'] = pd.to_datetime(df.iloc[:, final_date_idx], errors='coerce') if final_date_idx else pd.Timestamp.max
        clean_df['MaturityDate'] = pd.to_datetime(df.iloc[:, maturity_date_idx], errors='coerce') if maturity_date_idx else pd.NaT
        
        def calc_tenure(row):
            if pd.notna(row['MaturityDate']) and pd.notna(row['IssueDate']):
                days = (row['MaturityDate'] - row['IssueDate']).days
                return f"{int(round(days/30))}個月" 
            return "-"
        clean_df['Tenure'] = clean_df.apply(calc_tenure, axis=1)

        clean_df['KO_Pct'] = df.iloc[:, ko_idx].apply(clean_percentage)
        clean_df['KI_Pct'] = df.iloc[:, ki_idx].apply(clean_percentage)
        clean_df['Strike_Pct'] = df.iloc[:, strike_idx].apply(clean_percentage) if strike_idx else 100.0
        clean_df['KO_Type'] = df.iloc[:, ko_type_idx] if ko_type_idx else ""
        clean_df['KI_Type'] = df.iloc[:, ki_type_idx] if ki_type_idx else "AKI"

        # 讀取標的 (支援最多5支)
        # 關鍵邏輯：如果是新版格式(沒有進場價欄位)，我們要把 Initial 設為 0，稍後自動去抓
        for i in range(1, 6):
            if i == 1: tx_idx = t1_idx
            else:
                tx_idx, _ = find_col_index(cols, [f"標的{i}"])
                # 容錯：有時候是 標的1, 標的2... 有時候是 標的1, 標的1價格, 標的2...
                if tx_idx is None: 
                    # 猜測舊版格式 (標的佔2欄)
                    possible_idx = t1_idx + (i-1)*2
                    if possible_idx < len(df.columns): tx_idx = possible_idx
            
            if tx_idx is not None and tx_idx < len(df.columns):
                # 這裡做代號清洗
                raw_ticker = df.iloc[:, tx_idx]
                clean_df[f'T{i}_Code'] = raw_ticker.apply(clean_ticker_symbol)
                
                # 嘗試找進場價 (舊版)
                if tx_idx + 1 < len(df.columns):
                    # 檢查下一欄是否為數字 (進場價)
                    sample_val = df.iloc[0, tx_idx+1]
                    try:
                        float(sample_val) # 如果可以轉數字，當作是進場價
                        clean_df[f'T{i}_Initial'] = pd.to_numeric(df.iloc[:, tx_idx + 1], errors='coerce').fillna(0)
                    except:
                        # 不能轉數字，代表下一欄可能是別的東西 (新版格式)，初始價設為 0 (等等自動抓)
                        clean_df[f'T{i}_Initial'] = 0
                else:
                    clean_df[f'T{i}_Initial'] = 0
            else:
                clean_df[f'T{i}_Code'] = ""
                clean_df[f'T{i}_Initial'] = 0

        clean_df = clean_df.dropna(subset=['ID'])

        # 4. 準備下載資料
        today_ts = pd.Timestamp(real_today)
        min_trade_date = clean_df['TradeDate'].min()
        
        # 為了抓進場價，開始時間要涵蓋最早的交易日
        if pd.isna(min_trade_date):
            start_download_date = today_ts - timedelta(days=30)
        else:
            start_download_date = min_trade_date - timedelta(days=7) # 多抓一週緩衝

        all_tickers = []
        for i in range(1, 6):
            if f'T{i}_Code' in clean_df.columns:
                ts = clean_df[f'T{i}_Code'].dropna().unique().tolist()
                all_tickers.extend([t for t in ts if t != ""])
        all_tickers = list(set(all_tickers))

        if not all_tickers:
            st.error("❌ 找不到任何有效的標的代號。")
            st.stop()

        st.info(f"⏳ 正在下載美股資料... (涵蓋範圍: {start_download_date.strftime('%Y-%m-%d')} ~ 今日)")
        
        try:
            # 一次下載所有歷史資料
            history_data = yf.download(all_tickers, start=start_download_date, end=today_ts + timedelta(days=1))['Close']
        except Exception as e:
            st.error(f"美股連線失敗: {e}")
            st.stop()

        # 5. 核心運算
        results = []
        individual_messages = [] 
        admin_summary_list = []
        lookback_date = today_ts - timedelta(days=lookback_days)

        for index, row in clean_df.iterrows():
            # 參數設定
            ko_thresh_val = row['KO_Pct'] if pd.notna(row['KO_Pct']) else 100.0
            ki_thresh_val = row['KI_Pct'] if pd.notna(row['KI_Pct']) else 60.0
            strike_thresh_val = row['Strike_Pct'] if pd.notna(row['Strike_Pct']) else 100.0
            
            ko_thresh = ko_thresh_val / 100.0
            ki_thresh = ki_thresh_val / 100.0
            strike_thresh = strike_thresh_val / 100.0
            nc_months = parse_nc_months(row['KO_Type'])
            nc_end_date = row['IssueDate'] + relativedelta(months=nc_months)
            
            assets = []
            
            # --- 處理每一個標的 (包含自動補抓進場價) ---
            for i in range(1, 6):
                code = row.get(f'T{i}_Code', "")
                if code == "": continue
                
                initial = float(row.get(f'T{i}_Initial', 0))
                
                # 🌟 如果 Excel 沒填進場價 (==0)，則自動去抓交易日那天的收盤價
                if initial == 0:
                    trade_date = row['TradeDate']
                    if pd.notna(trade_date):
                        try:
                            # 嘗試抓取交易日當天
                            if len(all_tickers) == 1: s = history_data
                            else: s = history_data[code]
                            
                            # 抓取該日期 (如果當天沒開盤，往後找最近的一天)
                            # 使用 asof 或 reindex 比較複雜，這裡用簡單的 slice
                            price_on_trade = s[s.index >= trade_date].head(1)
                            if not price_on_trade.empty:
                                initial = float(price_on_trade.iloc[0])
                        except:
                            initial = 0 # 抓不到
                
                if initial > 0:
                    assets.append({
                        'code': code, 
                        'initial': initial, 
                        'strike_price': initial * strike_thresh, 
                        'locked_ko': False, 
                        'hit_ki': False, 
                        'perf': 0.0, 
                        'price': 0.0,
                        'ko_record': '',
                        'ki_record': ''
                    })
            
            if not assets: continue

            # --- 取得最新報價與表現 ---
            for asset in assets:
                try:
                    if len(all_tickers) == 1: s = history_data
                    else: s = history_data[asset['code']]
                    
                    # 抓最近一筆收盤價
                    valid_s = s[s.index <= today_ts].dropna()
                    if not valid_s.empty:
                        curr = float(valid_s.iloc[-1])
                        asset['price'] = curr
                        asset['perf'] = curr / asset['initial']
                except: asset['price'] = 0

            # --- 回測 (判斷 KO/KI) ---
            product_status = "Running"
            early_redemption_date = None
            is_aki = "AKI" in str(row['KI_Type']).upper()

            # 只有當已經發行後才開始回測
            if row['IssueDate'] <= today_ts:
                # 取得發行日到今天的數據
                backtest_data = history_data[(history_data.index >= row['IssueDate']) & (history_data.index <= today_ts)]
                
                if not backtest_data.empty:
                    for date, prices in backtest_data.iterrows():
                        if product_status == "Early Redemption": break
                        
                        is_post_nc = date >= nc_end_date
                        all_locked = True
                        
                        for asset in assets:
                            # 取得當日價格
                            try:
                                if len(all_tickers) == 1: price = float(prices)
                                else: price = float(prices[asset['code']])
                            except: price = float('nan')
                            
                            if pd.isna(price) or price == 0:
                                if not asset['locked_ko']: all_locked = False
                                continue
                            
                            perf = price / asset['initial']
                            date_str = date.strftime('%Y/%m/%d')
                            
                            # 檢查 KI (AKI: 每天觀察)
                            if is_aki and perf < ki_thresh:
                                if not asset['hit_ki']:
                                    asset['hit_ki'] = True
                                    asset['ki_record'] = f"@{price:.2f} ({date_str})"
                            
                            # 檢查 KO (過了 NC 且每天觀察 Daily)
                            # 注意：這裡簡化假設是 Daily Memory。如果是 Monthly 需另外判斷日期。
                            if not asset['locked_ko']:
                                if is_post_nc and perf >= ko_thresh:
                                    asset['locked_ko'] = True 
                                    asset['ko_record'] = f"@{price:.2f} ({date_str})"
                            
                            if not asset['locked_ko']: all_locked = False
                        
                        # 如果當天所有標的都 Lock KO -> 出場
                        if all_locked:
                            product_status = "Early Redemption"
                            early_redemption_date = date

            # --- 整理輸出資訊 ---
            locked_list = []; waiting_list = []; hit_ki_list = []; shadow_ko_list = []
            detail_cols = {}
            asset_detail_str = "" 

            for i, asset in enumerate(assets):
                # EKI 判斷 (到期當天 KI)
                if asset['price'] > 0:
                    if not is_aki and asset['perf'] < ki_thresh: 
                        asset['hit_ki'] = True # 暫時標記為破 KI (如果是到期日會真的算破)
                    if asset['perf'] >= ko_thresh and not asset['locked_ko']:
                        shadow_ko_list.append(asset['code'])

                if asset['locked_ko']: locked_list.append(asset['code'])
                else: waiting_list.append(asset['code'])
                if asset['hit_ki']: hit_ki_list.append(asset['code'])
                
                p_pct = round(asset['perf']*100, 2) if asset['price'] > 0 else 0.0
                status_icon = "✅" if asset['locked_ko'] else "⚠️" if asset['hit_ki'] else ""
                
                # 顯示資訊：代號 / 進場價 / 現價
                price_display = round(asset['price'], 2) if asset['price'] > 0 else "N/A"
                initial_display = round(asset['initial'], 2)
                
                cell_text = f"【{asset['code']}】\n原: {initial_display}\n現: {price_display}\n({p_pct}%) {status_icon}"
                if asset['locked_ko']: cell_text += f"\nKO {asset['ko_record']}"
                if asset['hit_ki']: cell_text += f"\nKI {asset['ki_record']}"
                detail_cols[f"T{i+1}_Detail"] = cell_text
                
                asset_detail_str += f"{asset['code']}: {p_pct}% {status_icon} (原:{initial_display})\n"

            hit_any_ki = any(a['hit_ki'] for a in assets)
            all_above_strike_now = all((a['perf'] >= strike_thresh if a['price'] > 0 else False) for a in assets)
            
            valid_assets = [a for a in assets if a['perf'] > 0]
            if valid_assets:
                worst_asset = min(valid_assets, key=lambda x: x['perf'])
                worst_perf = worst_asset['perf']
                worst_code = worst_asset['code']
            else:
                worst_perf = 0; worst_code = "N/A"
            
            final_status = ""
            line_status_short = "" 
            need_notify = False

            # 狀態判定邏輯
            if today_ts < row['IssueDate']:
                final_status = "⏳ 未發行"
            elif product_status == "Early Redemption":
                final_status = f"🎉 提前出場\n({early_redemption_date.strftime('%Y-%m-%d')})"
                # 檢查是否為「新」事件
                if early_redemption_date >= lookback_date:
                    line_status_short = "🎉 恭喜！已提前出場 (KO)"
                    need_notify = True
                else:
                    line_status_short = f"🎉 已於 {early_redemption_date.strftime('%Y-%m-%d')} 提前出場 (舊)"
                    need_notify = False
            elif pd.notna(row['ValuationDate']) and today_ts >= row['ValuationDate']:
                # 到期
                is_recent = row['ValuationDate'] >= lookback_date
                if all_above_strike_now:
                     final_status = "💰 到期獲利"
                     line_status_short = "💰 到期獲利"
                elif hit_any_ki:
                     final_status = f"😭 到期接股"
                     line_status_short = f"😭 到期接股 (Worst: {worst_code})"
                else:
                     final_status = "🛡️ 到期保本"
                     line_status_short = "🛡️ 到期保本"
                
                need_notify = is_recent
                if not is_recent: line_status_short += " (舊)"
            else:
                # 執行中
                if today_ts < nc_end_date:
                    final_status = f"🔒 NC閉鎖期\n(至 {nc_end_date.strftime('%Y-%m-%d')})"
                else:
                    wait_str = ",".join(waiting_list) if waiting_list else "無"
                    final_status = f"👀 比價中"
                
                if hit_any_ki:
                    final_status += f"\n⚠️ KI已破: {','.join(hit_ki_list)}"
                    line_status_short = f"⚠️ 注意：KI 已跌破 ({','.join(hit_ki_list)})"
                    need_notify = notify_ki_daily

            if line_status_short:
                admin_summary_list.append(f"● {row['ID']} ({row['Name']}): {line_status_short}")

            # 收集發送名單
            line_ids = [x.strip() for x in re.split(r'[;,，]', str(row.get('Line_ID', ''))) if x.strip()]
            emails = [x.strip() for x in re.split(r'[;,，]', str(row.get('Email', ''))) if x.strip()]
            
            common_msg_body = (
                f"Hi {row['Name']} 您好，\n"
                f"您的結構型商品 {row['ID']} 最新狀態：\n\n"
                f"【{line_status_short}】\n\n"
                f"{asset_detail_str}"
                f"📅 到期日: {row['MaturityDate'].strftime('%Y-%m-%d') if pd.notna(row['MaturityDate']) else '-'}\n"
                f"------------------\n"
                f"貼心通知"
            )

            if need_notify and line_status_short:
                # LINE
                for uid in line_ids:
                    if uid.startswith("U") or uid.startswith("C"):
                        individual_messages.append({'type': 'line', 'target': uid, 'msg': common_msg_body})
                
                # Email
                for mail in emails:
                    if "@" in mail:
                        subject = f"【ELN通知】{row['ID']} 最新狀態"
                        mail_body = common_msg_body + "\n(本信件由系統自動發送)"
                        individual_messages.append({'type': 'email', 'target': mail, 'subj': subject, 'msg': mail_body})

            # 收集結果到 DataFrame
            row_res = {
                "債券代號": row['ID'], "Name": row['Name'],
                "狀態": final_status, "最差表現": f"{round(worst_perf*100, 2)}%",
                "交易日": row['TradeDate'].strftime('%Y-%m-%d') if pd.notna(row['TradeDate']) else "-"
            }
            row_res.update(detail_cols)
            results.append(row_res)

        # 6. 顯示與操作
        if not results:
            st.warning("⚠️ 無資料")
        else:
            final_df = pd.DataFrame(results)
            
            # 設定顏色
            def color_status(val):
                if "提前" in str(val) or "獲利" in str(val): return 'background-color: #d4edda; color: green'
                if "接股" in str(val) or "KI" in str(val): return 'background-color: #f8d7da; color: red'
                if "未發行" in str(val) or "NC" in str(val): return 'background-color: #fff3cd; color: #856404'
                return ''

            t_cols = [c for c in final_df.columns if '_Detail' in c]; t_cols.sort()
            display_cols = ['債券代號', 'Name', '狀態', '最差表現'] + t_cols + ['交易日']
            
            st.subheader("📋 監控列表")
            st.dataframe(final_df[display_cols].style.applymap(color_status, subset=['狀態']), height=600, use_container_width=True)

            st.markdown("### 📢 發送操作")
            
            if st.session_state['is_sent']:
                st.success("✅ 發送完成！")
                if st.button("🔄 重置"):
                    st.session_state['is_sent'] = False
                    st.rerun()
            else:
                count = len(individual_messages)
                btn_label = f"🚀 發送通知 (預計: {count} 則)"
                
                if st.button(btn_label, type="primary"):
                    success_cnt = 0
                    bar = st.progress(0, text="發送中...")
                    
                    for idx, item in enumerate(individual_messages):
                        res = False
                        if item['type'] == 'line':
                            res = send_line_push(item['target'], item['msg'])
                        elif item['type'] == 'email':
                            res = send_email_gmail(item['target'], item['subj'], item['msg'])
                        
                        if res: success_cnt += 1
                        bar.progress((idx+1)/count)
                    
                    bar.empty()
                    
                    # 發送給管理員 (LINE)
                    if admin_summary_list and MY_LINE_USER_ID:
                        summary = "【ELN 戰情快報】\n" + "\n".join(admin_summary_list)
                        send_line_push(MY_LINE_USER_ID, summary)

                    st.session_state['is_sent'] = True
                    st.success(f"🎉 成功發送 {success_cnt} 則通知！")
                    st.balloons()

    except Exception as e:
        st.error(f"發生錯誤：{e}")
