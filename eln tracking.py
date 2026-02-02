import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import re
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from dateutil.relativedelta import relativedelta

# --- 設定網頁 ---
st.set_page_config(page_title="ELN 智能戰情室 (智能樣式版)", layout="wide")

# ==========================================
# 📧 Email 伺服器設定
# ==========================================
try:
    SMTP_SERVER = st.secrets.get("SMTP_SERVER", "smtp.gmail.com")
    SMTP_PORT = st.secrets.get("SMTP_PORT", 587)
    SENDER_EMAIL = st.secrets.get("SENDER_EMAIL", "")
    SENDER_PASSWORD = st.secrets.get("SENDER_PASSWORD", "")
except Exception:
    SMTP_SERVER = ""
    SMTP_PORT = 587
    SENDER_EMAIL = ""
    SENDER_PASSWORD = ""

# ==========================================
# 🔄 狀態初始化
# ==========================================
if 'last_processed_file' not in st.session_state:
    st.session_state['last_processed_file'] = None
if 'is_sent' not in st.session_state:
    st.session_state['is_sent'] = False

# --- 側邊欄 ---
with st.sidebar:
    st.header("📧 設定中心")
    if SENDER_EMAIL and SENDER_PASSWORD:
        st.success(f"✅ 寄件者已設定：\n{SENDER_EMAIL}")
    else:
        st.error("❌ Email 尚未設定 Secrets")

    st.markdown("---")
    real_today = datetime.now()
    st.info(f"📅 今天日期：{real_today.strftime('%Y-%m-%d')}")
    
    st.markdown("---")
    st.header("🔔 通知過濾")
    lookback_days = st.slider("顯示幾天內發生的事件？", min_value=1, max_value=30, value=3)
    st.info("💡 **樣式說明**：\n已提前出場超過 **2天** 的商品，將會自動顯示為~~刪除線~~並變灰，方便您聚焦在即將到期的案件。")

# --- 函數區 ---

def clean_ticker_symbol(ticker):
    if pd.isna(ticker): return ""
    t = str(ticker).strip().upper()
    t = re.sub(r'\s+(UW|UN|UQ|UP|US)$', '', t)
    if t.endswith(" JT"): return t.replace(" JT", ".T") 
    if t.endswith(" TT"): return t.replace(" TT", ".TW") 
    if t.endswith(" HK"): return t.replace(" HK", ".HK") 
    return t

def send_email_html(to_email, subject, html_content):
    if not SENDER_EMAIL or not SENDER_PASSWORD: return False
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = Header(subject, 'utf-8')
        msg.attach(MIMEText(html_content, 'html', 'utf-8'))

        server = smtplib.SMTP(SMTP_SERVER, int(SMTP_PORT))
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"Email 發送失敗 ({to_email}): {e}")
        return False

def parse_ko_settings(ko_price_val):
    s = str(ko_price_val).strip()
    initial_ko = 100.0
    step_rate = 0.0
    if pd.isna(ko_price_val) or s == "": return initial_ko, step_rate
    match = re.search(r'^(\d+(?:\.\d+)?)', s)
    if match: initial_ko = float(match.group(1))
    step_match = re.search(r'[\(（].*?(\d+(?:\.\d+)?)%?\s*(?:遞減|step|less|down)', s, re.IGNORECASE)
    if step_match: step_rate = float(step_match.group(1))
    return initial_ko, step_rate

def parse_nc_months(ko_type_val):
    s = str(ko_type_val).upper().strip()
    if pd.isna(ko_type_val) or s == "" or s == "NAN": return 1 
    match = re.search(r'(?:NC|LOCK|NON-CALL)\s*[:\-]?\s*(\d+)', s)
    if match: return int(match.group(1))
    return 1

def is_period_end_check(ko_type_val):
    s = str(ko_type_val).upper().strip()
    return "PERIOD END" in s or "MONTHLY" in s

def calculate_maturity(row, issue_date_col, tenure_col):
    if 'MaturityDate' in row and pd.notna(row['MaturityDate']): return row['MaturityDate']
    issue_date = row.get(issue_date_col)
    tenure_str = str(row.get(tenure_col, ""))
    if pd.isna(issue_date) or issue_date == pd.NaT: return pd.NaT
    try:
        months_to_add = 0
        match_m = re.search(r'(\d+)\s*M', tenure_str, re.IGNORECASE)
        match_y = re.search(r'(\d+)\s*Y', tenure_str, re.IGNORECASE)
        if match_m: months_to_add = int(match_m.group(1))
        elif match_y: months_to_add = int(match_y.group(1)) * 12
        elif tenure_str.isdigit(): months_to_add = int(tenure_str)
        if months_to_add > 0: return issue_date + relativedelta(months=months_to_add)
    except: pass
    return pd.NaT

def clean_percentage(val):
    if pd.isna(val) or str(val).strip() == "": return None
    try:
        s = str(val).replace('%', '').replace(',', '').strip()
        s = re.split(r'[\(（]', s)[0]
        return float(s)
    except: return None

def clean_name_str(val):
    if pd.isna(val): return "貴賓"
    s = str(val).strip()
    if s.lower() == 'nan' or s == "": return "貴賓"
    return s

def find_col_index(columns, include_keywords, exclude_keywords=None):
    for idx, col_name in enumerate(columns):
        col_str = str(col_name).strip().lower().replace(" ", "")
        if exclude_keywords:
            if any(ex in col_str for ex in exclude_keywords): continue
        if any(inc in col_str for inc in include_keywords):
            return idx, col_name
    return None, None

# --- 主畫面 ---
st.title("📧 ELN 智能戰情室 - 智能樣式版")

uploaded_file = st.file_uploader("請上傳 Excel", type=['xlsx', 'csv'], key="uploader")

if uploaded_file:
    if st.session_state['last_processed_file'] != uploaded_file.name:
        st.session_state['last_processed_file'] = uploaded_file.name
        st.session_state['is_sent'] = False

if uploaded_file is not None:
    try:
        try:
            df = pd.read_excel(uploaded_file, sheet_name=0, header=0, engine='openpyxl')
        except:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file)

        df = df.dropna(how='all')
        if df.iloc[0].astype(str).str.contains("進場價").any():
            df = df.iloc[1:].reset_index(drop=True)
        cols = df.columns.tolist()
        
        # 欄位定位
        id_idx, _ = find_col_index(cols, ["債券", "代號", "id", "商品代號"]) or (0, "")
        type_idx, _ = find_col_index(cols, ["商品類型", "ProductType", "type"], exclude_keywords=["ko", "ki"]) 
        strike_idx, _ = find_col_index(cols, ["strike", "執行", "履約"])
        ko_idx, _ = find_col_index(cols, ["ko", "提前"], exclude_keywords=["strike", "執行", "ki", "type"])
        ko_type_idx, _ = find_col_index(cols, ["ko類型", "kotype"]) or find_col_index(cols, ["類型", "type"], exclude_keywords=["ki", "ko", "商品"])
        ki_idx, _ = find_col_index(cols, ["ki", "下檔"], exclude_keywords=["ko", "type"])
        ki_type_idx, _ = find_col_index(cols, ["ki類型", "kitype"])
        t1_idx, _ = find_col_index(cols, ["標的1", "ticker1"])
        trade_date_idx, _ = find_col_index(cols, ["交易日"])
        issue_date_idx, _ = find_col_index(cols, ["發行日"])
        final_date_idx, _ = find_col_index(cols, ["最終", "評價"])
        maturity_date_idx, _ = find_col_index(cols, ["到期", "maturity"])
        tenure_idx, _ = find_col_index(cols, ["天期", "term", "tenure"])
        name_idx, _ = find_col_index(cols, ["理專", "姓名", "客戶"])
        email_idx, _ = find_col_index(cols, ["email", "e-mail", "mail", "信箱", "電子郵件"])

        if t1_idx is None:
            st.error("❌ 無法辨識「標的1」欄位，請檢查 Excel 表頭。")
            st.stop()

        clean_df = pd.DataFrame()
        clean_df['ID'] = df.iloc[:, id_idx]
        if name_idx is not None: clean_df['Name'] = df.iloc[:, name_idx].apply(clean_name_str)
        else: clean_df['Name'] = "貴賓"
        if email_idx is not None: clean_df['Email'] = df.iloc[:, email_idx].astype(str).replace('nan', '').str.strip()
        else: clean_df['Email'] = ""

        if type_idx is not None: clean_df['Product_Type'] = df.iloc[:, type_idx].astype(str).fillna("FCN")
        else: clean_df['Product_Type'] = "FCN"

        clean_df['TradeDate'] = pd.to_datetime(df.iloc[:, trade_date_idx], errors='coerce') if trade_date_idx else pd.NaT
        clean_df['IssueDate'] = pd.to_datetime(df.iloc[:, issue_date_idx], errors='coerce') if issue_date_idx else pd.Timestamp.min
        if maturity_date_idx: clean_df['MaturityDate'] = pd.to_datetime(df.iloc[:, maturity_date_idx], errors='coerce')
        else: clean_df['MaturityDate'] = pd.NaT
        clean_df['ValuationDate'] = pd.to_datetime(df.iloc[:, final_date_idx], errors='coerce') if final_date_idx else pd.NaT
        clean_df['TenureStr'] = df.iloc[:, tenure_idx] if tenure_idx else ""

        for idx, row in clean_df.iterrows():
            if pd.isna(row['MaturityDate']):
                calc_date = calculate_maturity(row, 'IssueDate', 'TenureStr')
                clean_df.at[idx, 'MaturityDate'] = calc_date
                if pd.isna(row['ValuationDate']): clean_df.at[idx, 'ValuationDate'] = calc_date

        def calc_tenure_display(row):
            if row['TenureStr'] != "": return str(row['TenureStr'])
            if pd.notna(row['MaturityDate']) and pd.notna(row['IssueDate']):
                days = (row['MaturityDate'] - row['IssueDate']).days
                return f"{int(round(days/30))}M" 
            return "-"
        clean_df['Tenure'] = clean_df.apply(calc_tenure_display, axis=1)

        clean_df['KO_Initial'], clean_df['KO_Step'] = zip(*df.iloc[:, ko_idx].apply(parse_ko_settings))
        clean_df['KI_Pct'] = df.iloc[:, ki_idx].apply(clean_percentage)
        clean_df['Strike_Pct'] = df.iloc[:, strike_idx].apply(clean_percentage) if strike_idx else 100.0
        clean_df['KO_Type'] = df.iloc[:, ko_type_idx] if ko_type_idx else "NC1" 
        clean_df['KI_Type'] = df.iloc[:, ki_type_idx] if ki_type_idx else "AKI"

        for i in range(1, 6):
            if i == 1: tx_idx = t1_idx
            else:
                tx_idx, _ = find_col_index(cols, [f"標的{i}"])
                if tx_idx is None: 
                    possible_idx = t1_idx + (i-1)*2
                    if possible_idx < len(df.columns): tx_idx = possible_idx
            if tx_idx is not None and tx_idx < len(df.columns):
                raw_ticker = df.iloc[:, tx_idx]
                clean_df[f'T{i}_Code'] = raw_ticker.apply(clean_ticker_symbol)
                if tx_idx + 1 < len(df.columns):
                    sample_val = df.iloc[0, tx_idx+1]
                    try:
                        float(sample_val)
                        clean_df[f'T{i}_Initial'] = pd.to_numeric(df.iloc[:, tx_idx + 1], errors='coerce').fillna(0)
                    except: clean_df[f'T{i}_Initial'] = 0
                else: clean_df[f'T{i}_Initial'] = 0
            else:
                clean_df[f'T{i}_Code'] = ""; clean_df[f'T{i}_Initial'] = 0

        clean_df = clean_df.dropna(subset=['ID'])

        today_ts = pd.Timestamp(real_today)
        min_trade_date = clean_df['TradeDate'].min()
        if pd.isna(min_trade_date): start_download_date = today_ts - timedelta(days=30)
        else: start_download_date = min_trade_date - timedelta(days=7)

        all_tickers = []
        for i in range(1, 6):
            if f'T{i}_Code' in clean_df.columns:
                ts = clean_df[f'T{i}_Code'].dropna().unique().tolist()
                all_tickers.extend([t for t in ts if t != ""])
        all_tickers = list(set(all_tickers))

        if not all_tickers:
            st.error("❌ 找不到有效的標的代號。")
            st.stop()

        st.info(f"⏳ 下載美股資料... ({start_download_date.strftime('%Y-%m-%d')} ~ 今日)")
        try:
            history_data = yf.download(all_tickers, start=start_download_date, end=today_ts + timedelta(days=1))['Close']
        except Exception as e:
            st.error(f"美股連線失敗: {e}")
            st.stop()

        results = []
        individual_messages = [] 
        lookback_date = today_ts - timedelta(days=lookback_days)

        for index, row in clean_df.iterrows():
            ki_thresh_val = row['KI_Pct'] if pd.notna(row['KI_Pct']) else 60.0
            strike_thresh_val = row['Strike_Pct'] if pd.notna(row['Strike_Pct']) else 100.0
            ko_initial_val = row['KO_Initial']
            ko_step_val = row['KO_Step']
            ki_thresh = ki_thresh_val / 100.0
            strike_thresh = strike_thresh_val / 100.0
            nc_months = parse_nc_months(row['KO_Type'])
            nc_end_date = row['IssueDate'] + relativedelta(months=nc_months)
            
            is_dra = "DRA" in str(row['Product_Type']).upper()
            is_period_end = is_period_end_check(row['KO_Type'])
            is_aki = "AKI" in str(row['KI_Type']).upper()
            
            assets = []
            for i in range(1, 6):
                code = row.get(f'T{i}_Code', "")
                if code == "": continue
                initial = float(row.get(f'T{i}_Initial', 0))
                if initial == 0:
                    trade_date = row['TradeDate']
                    if pd.notna(trade_date):
                        try:
                            if len(all_tickers) == 1: s = history_data
                            else: s = history_data[code]
                            price_on_trade = s[s.index >= trade_date].head(1)
                            if not price_on_trade.empty: initial = float(price_on_trade.iloc[0])
                        except: initial = 0
                if initial > 0:
                    assets.append({
                        'code': code, 'initial': initial, 'strike_price': initial * strike_thresh, 
                        'locked_ko': False, 'hit_ki': False, 'perf': 0.0, 'price': 0.0, 
                        'ko_record': '', 'ki_record': '',
                        'eki_risk': False
                    })
            if not assets: continue

            for asset in assets:
                try:
                    if len(all_tickers) == 1: s = history_data
                    else: s = history_data[asset['code']]
                    valid_s = s[s.index <= today_ts].dropna()
                    if not valid_s.empty:
                        curr = float(valid_s.iloc[-1])
                        asset['price'] = curr
                        asset['perf'] = curr / asset['initial']
                except: asset['price'] = 0

            months_passed = 0
            if pd.notna(row['IssueDate']):
                months_passed = (today_ts.year - row['IssueDate'].year) * 12 + today_ts.month - row['IssueDate'].month
                if months_passed < 0: months_passed = 0
            current_ko_pct = ko_initial_val - (ko_step_val * months_passed)
            current_ko_thresh = current_ko_pct / 100.0

            product_status = "Running"
            early_redemption_date = None
            
            if row['IssueDate'] <= today_ts:
                backtest_data = history_data[(history_data.index >= row['IssueDate']) & (history_data.index <= today_ts)]
                if not backtest_data.empty:
                    for date, prices in backtest_data.iterrows():
                        if product_status == "Early Redemption": break
                        is_post_nc = date >= nc_end_date
                        is_obs_date = True
                        if is_period_end:
                            if date.day != row['IssueDate'].day: is_obs_date = False
                        
                        m_pass = (date.year - row['IssueDate'].year) * 12 + date.month - row['IssueDate'].month
                        if date.day < row['IssueDate'].day: m_pass -= 1
                        if m_pass < 0: m_pass = 0
                        day_ko_val = ko_initial_val - (ko_step_val * m_pass)
                        day_ko_thresh = day_ko_val / 100.0

                        all_locked = True
                        for asset in assets:
                            try:
                                if len(all_tickers) == 1: price = float(prices)
                                else: price = float(prices[asset['code']])
                            except: price = float('nan')
                            if pd.isna(price) or price == 0:
                                if not asset['locked_ko']: all_locked = False
                                continue
                            
                            perf = price / asset['initial']
                            date_str = date.strftime('%Y/%m/%d')
                            if is_aki and perf < ki_thresh and not asset['hit_ki']:
                                asset['hit_ki'] = True
                                asset['ki_record'] = f"@{price:.2f} ({date_str})"
                            if not asset['locked_ko']:
                                if is_post_nc:
                                    if is_period_end and not is_obs_date: pass
                                    else:
                                        if perf >= day_ko_thresh:
                                            asset['locked_ko'] = True 
                                            asset['ko_record'] = f"@{price:.2f} ({date_str})"
                            if not asset['locked_ko']: all_locked = False
                        
                        if all_locked:
                            product_status = "Early Redemption"
                            early_redemption_date = date

            locked_list = []; waiting_list = []; hit_ki_list = []
            detail_cols = {}
            asset_rows_html = ""
            any_below_strike_today = False
            dra_fail_list = []
            any_eki_risk_today = False

            for i, asset in enumerate(assets):
                if asset['price'] > 0:
                    if is_aki:
                        if asset['perf'] < ki_thresh: asset['hit_ki'] = True 
                    else:
                        if asset['perf'] < ki_thresh: 
                            asset['eki_risk'] = True
                            any_eki_risk_today = True

                    if is_dra and asset['perf'] < strike_thresh:
                        any_below_strike_today = True
                        dra_fail_list.append(asset['code'])

                if asset['locked_ko']: locked_list.append(asset['code'])
                else: waiting_list.append(asset['code'])
                if asset['hit_ki']: hit_ki_list.append(asset['code'])
                
                p_pct = round(asset['perf']*100, 2) if asset['price'] > 0 else 0.0
                status_icon = "✅" if asset['locked_ko'] else "⚠️" if asset['hit_ki'] else ""
                if asset['eki_risk']: status_icon = "📉"
                if is_dra and asset['price'] > 0:
                    if asset['perf'] < strike_thresh: status_icon += "🛑無息"
                    else: status_icon += "💸"

                price_display = round(asset['price'], 2) if asset['price'] > 0 else "N/A"
                initial_display = round(asset['initial'], 2)
                
                cell_text = f"【{asset['code']}】\n原: {initial_display}\n現: {price_display}\n({p_pct}%) {status_icon}"
                if asset['locked_ko']: cell_text += f"\nKO {asset['ko_record']}"
                if asset['hit_ki']: cell_text += f"\nKI {asset['ki_record']}"
                detail_cols[f"T{i+1}_Detail"] = cell_text

                # 📧 HTML 表格
                row_style = ""
                if asset['hit_ki'] or asset['eki_risk'] or (is_dra and asset['perf'] < strike_thresh):
                    row_style = "color: red; font-weight: bold;"
                elif asset['locked_ko']:
                    row_style = "color: green; font-weight: bold;"
                
                asset_rows_html += f"""
                <tr>
                    <td style="border: 1px solid #ddd; padding: 8px;">{asset['code']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px;">{initial_display}</td>
                    <td style="border: 1px solid #ddd; padding: 8px;">{price_display}</td>
                    <td style="border: 1px solid #ddd; padding: 8px; {row_style}">{p_pct}% {status_icon}</td>
                </tr>
                """

            hit_any_ki = any(a['hit_ki'] for a in assets)
            all_above_strike_now = all((a['perf'] >= strike_thresh if a['price'] > 0 else False) for a in assets)
            valid_assets = [a for a in assets if a['perf'] > 0]
            if valid_assets: worst_perf = min(valid_assets, key=lambda x: x['perf'])['perf']
            else: worst_perf = 0
            
            status_msgs = []
            line_status_short = ""
            need_notify = False
            title_color = "#333333"

            if today_ts < row['IssueDate']:
                status_msgs.append("⏳ 未發行")
            elif product_status == "Early Redemption":
                status_msgs.append(f"🎉 提前出場 ({early_redemption_date.strftime('%Y-%m-%d')})")
                if early_redemption_date >= lookback_date:
                    line_status_short = "🎉 恭喜！已提前出場 (KO)"
                    title_color = "green"
                    need_notify = True
            elif pd.notna(row['ValuationDate']) and today_ts >= row['ValuationDate']:
                final_hit_ki = False
                for a in assets:
                     if a['perf'] < ki_thresh: final_hit_ki = True
                if all_above_strike_now:
                     status_msgs.append("💰 到期獲利")
                     line_status_short = "💰 到期獲利"
                     title_color = "green"
                elif final_hit_ki:
                     status_msgs.append("😭 到期接股")
                     line_status_short = "😭 到期接股"
                     title_color = "red"
                else:
                     status_msgs.append("🛡️ 到期保本")
                     line_status_short = "🛡️ 到期保本"
                if row['ValuationDate'] >= lookback_date: need_notify = True
            else:
                if today_ts < nc_end_date:
                    status_msgs.append(f"🔒 NC閉鎖期 (至 {nc_end_date.strftime('%Y-%m-%d')})")
                else:
                    if is_period_end: status_msgs.append(f"👀 比價中 (月月比)")
                    else: status_msgs.append("👀 比價中 (Daily)")
                
                if ko_step_val > 0: status_msgs.append(f"📉 目前KO門檻: {current_ko_pct}%")

                if hit_any_ki:
                    status_msgs.insert(0, f"☠️ 已跌破KI ({','.join(hit_ki_list)})")
                    line_status_short = f"⚠️ 警告：已跌破 KI ({','.join(hit_ki_list)})"
                    title_color = "red"
                    need_notify = True 
                elif any_eki_risk_today:
                     status_msgs.insert(0, f"📉 市價低於KI (EKI觀察中)")

                if is_dra:
                    if any_below_strike_today:
                        status_msgs.append(f"🛑 DRA暫停計息 ({','.join(dra_fail_list)})")
                        status_msgs.insert(0, f"⚠️ DRA 暫停計息")
                        need_notify = True

            final_status = "\n".join(status_msgs)
            target_email = row.get('Email', '')
            mat_date_str = row['MaturityDate'].strftime('%Y-%m-%d') if pd.notna(row['MaturityDate']) else "-"
            
            email_subject = f"【ELN通知】{row['ID']} - {line_status_short}" if line_status_short else f"【ELN週報】{row['ID']} 狀態報告"
            email_html_body = f"""
            <html>
            <body>
                <h3>Hi {row['Name']} 您好，</h3>
                <p>您的結構型商品 <b>{row['ID']}</b> ({row['Product_Type']}) 最新狀態如下：</p>
                <h2 style="color: {title_color};">{line_status_short}</h2>
                <p><b>詳細標的表現：</b></p>
                <table style="border-collapse: collapse; width: 100%;">
                    <tr style="background-color: #f2f2f2;">
                        <th style="border: 1px solid #ddd; padding: 8px;">代號</th>
                        <th style="border: 1px solid #ddd; padding: 8px;">進場價</th>
                        <th style="border: 1px solid #ddd; padding: 8px;">現價</th>
                        <th style="border: 1px solid #ddd; padding: 8px;">表現</th>
                    </tr>
                    {asset_rows_html}
                </table>
                <br>
                <p>📅 <b>到期日：</b> {mat_date_str}</p>
                <hr>
                <p style="font-size: 12px; color: gray;">此郵件為系統自動發送。</p>
            </body>
            </html>
            """

            if need_notify and line_status_short and target_email and "@" in str(target_email):
                individual_messages.append({
                    'target': target_email, 
                    'subject': email_subject, 
                    'html': email_html_body
                })

            row_res = {
                "債券代號": row['ID'], "Name": row['Name'], "Type": row['Product_Type'],
                "狀態": final_status, "最差表現": f"{round(worst_perf*100, 2)}%",
                "交易日": row['TradeDate'].strftime('%Y-%m-%d') if pd.notna(row['TradeDate']) else "-",
                "NC月份": f"{nc_months}M",
                "KO設定": f"{ko_initial_val}% (-{ko_step_val}%)" if ko_step_val > 0 else f"{ko_initial_val}%"
            }
            row_res.update(detail_cols)
            results.append(row_res)

        # --- 顯示與操作 ---
        if not results:
            st.warning("⚠️ 無資料")
        else:
            final_df = pd.DataFrame(results)
            
            # 🔥 新增：智能樣式函數
            def highlight_rows(row):
                styles = [''] * len(row)
                status = str(row['狀態'])
                
                # 1. 判斷是否為「久遠的 KO」
                is_old_ko = False
                if "提前出場" in status:
                    # 抓日期
                    match = re.search(r'(\d{4}-\d{2}-\d{2})', status)
                    if match:
                        try:
                            ko_date = datetime.strptime(match.group(1), '%Y-%m-%d')
                            days_diff = (datetime.now() - ko_date).days
                            if days_diff > 2:
                                is_old_ko = True
                        except: pass
                
                # 2. 舊 KO 刪除線樣式
                if is_old_ko:
                    return ['text-decoration: line-through; color: #cccccc;'] * len(row)
                    
                # 3. 一般樣式 (針對「狀態」欄位)
                for i, col in enumerate(row.index):
                    if col == '狀態':
                        if "跌破KI" in status or "接股" in status:
                            styles[i] = 'background-color: #f8d7da; color: red; font-weight: bold'
                        elif "EKI" in status or "暫停" in status:
                            styles[i] = 'background-color: #fff3cd; color: #856404'
                        elif "提前" in status or "獲利" in status:
                            styles[i] = 'background-color: #d4edda; color: green'
                return styles

            t_cols = [c for c in final_df.columns if '_Detail' in c]; t_cols.sort()
            display_cols = ['債券代號', 'Type', 'Name', '狀態', 'KO設定', '最差表現'] + t_cols + ['交易日']
            
            st.subheader("📋 監控列表")
            # 套用整行樣式
            st.dataframe(final_df[display_cols].style.apply(highlight_rows, axis=1), height=600, use_container_width=True)

            st.markdown("### 📧 信件發送操作")
            
            secrets_ok = (SENDER_EMAIL != "" and SENDER_PASSWORD != "")
            if not secrets_ok:
                st.error("⚠️ 未設定 Email，無法發送。請至 Secrets 設定。")
            
            if st.session_state['is_sent']:
                st.success("✅ 發送完成！")
                if st.button("🔄 重置"):
                    st.session_state['is_sent'] = False
                    st.rerun()
            else:
                count = len(individual_messages)
                btn_label = f"🚀 發送 Email 通知 (預計: {count} 封)"
                
                if st.button(btn_label, type="primary", disabled=not secrets_ok):
                    if count == 0:
                        st.warning("沒有需要通知的事件。")
                    else:
                        success_cnt = 0
                        bar = st.progress(0, text="正在發送信件...")
                        for idx, item in enumerate(individual_messages):
                            if send_email_html(item['target'], item['subject'], item['html']):
                                success_cnt += 1
                            bar.progress((idx+1)/count)
                        bar.empty()
                        st.session_state['is_sent'] = True
                        st.success(f"🎉 成功寄出 {success_cnt} 封 Email！")
                        st.balloons()

    except Exception as e:
        st.error(f"發生錯誤：{e}")
