import streamlit as st
import pandas as pd
import yfinance as yf
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import re
from dateutil.relativedelta import relativedelta

# --- 設定網頁 ---
st.set_page_config(page_title="ELN 專業監控戰情室", layout="wide")

# --- 側邊欄：設定 ---
with st.sidebar:
    st.header("📧 設定中心")
    sender_email = st.text_input("寄件人 Gmail", placeholder="example@gmail.com")
    sender_password = st.text_input("應用程式密碼", type="password", placeholder="16位數密碼")
    st.markdown("---")
    st.info("💡 **介面優化：**\n1. 狀態欄顯示：已鎖定/等待中標的\n2. 合併標的明細，解決表格過寬問題\n3. 標的代碼清楚列出")

# --- 函數：發送 Email ---
def send_email(sender, password, receiver, subject, body):
    if not sender or not password or not receiver:
        st.warning("⚠️ 寄件資料不完整")
        return False
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = receiver
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender, password)
        server.send_message(msg)
        server.quit()
        st.toast(f"✅ 已寄信給 {receiver}", icon="📩")
        return True
    except Exception as e:
        st.error(f"❌ 發送失敗：{e}")
        return False

# --- 函數：解析 NC 月份 ---
def parse_nc_months(ko_type_str):
    if pd.isna(ko_type_str) or str(ko_type_str).strip() == "":
        return 1 
    match = re.search(r'NC(\d+)', str(ko_type_str), re.IGNORECASE)
    if match:
        return int(match.group(1))
    return 1 

# --- 函數：數據清洗 ---
def clean_percentage(val):
    if pd.isna(val) or str(val).strip() == "":
        return None
    try:
        s = str(val).replace('%', '').replace(',', '').strip()
        return float(s)
    except:
        return None

# --- 函數：嚴格尋找欄位 ---
def find_col_index(columns, include_keywords, exclude_keywords=None):
    for idx, col_name in enumerate(columns):
        col_str = str(col_name).strip().lower()
        if exclude_keywords:
            if any(ex in col_str for ex in exclude_keywords): continue
        if any(inc in col_str for inc in include_keywords):
            return idx, col_name
    return None, None

# --- 主畫面 ---
st.title("📊 ELN 結構型商品 - 專業監控戰情室")
st.markdown("### 🚀 智能狀態摘要與版面瘦身版")

uploaded_file = st.file_uploader("請上傳 Excel (工作表1格式)", type=['xlsx'])

if uploaded_file is not None:
    try:
        # 1. 讀取與清洗
        try:
            df = pd.read_excel(uploaded_file, sheet_name=0, header=0, engine='openpyxl')
        except:
            df = pd.read_excel(uploaded_file, sheet_name=0, header=0)

        if df.iloc[0].astype(str).str.contains("進場價").any():
            df = df.iloc[1:].reset_index(drop=True)

        cols = df.columns.tolist()
        
        # --- 2. 欄位定位 ---
        id_idx, _ = find_col_index(cols, ["債券", "代號", "id"])
        if id_idx is None: id_idx = 0
        
        strike_idx, _ = find_col_index(cols, ["strike", "執行", "履約", "conversion"])
        ko_idx, _ = find_col_index(cols, ["ko", "knock-out", "提前", "autocall"], exclude_keywords=["strike", "執行", "履約", "ki", "type", "類型"])
        ko_type_idx, _ = find_col_index(cols, ["ko類型", "ko type", "autocall type"])
        if ko_type_idx is None: ko_type_idx, _ = find_col_index(cols, ["類型", "type"], exclude_keywords=["ki", "ko"])

        ki_idx, _ = find_col_index(cols, ["ki", "knock-in", "下檔", "barrier"], exclude_keywords=["ko", "type", "類型"])
        ki_type_idx, _ = find_col_index(cols, ["ki類型", "ki type"])
        
        t1_idx, _ = find_col_index(cols, ["標的1", "ticker 1"])
        
        issue_date_idx, _ = find_col_index(cols, ["發行日", "trade date", "start"])
        final_date_idx, _ = find_col_index(cols, ["最終", "評價", "final", "valuation"])
        email_idx, _ = find_col_index(cols, ["email", "信箱"])
        name_idx, _ = find_col_index(cols, ["理專", "姓名", "客戶"])

        if t1_idx is None or ko_idx is None:
            st.error("❌ 嚴重錯誤：無法辨識關鍵欄位 (KO 或 標的1)。")
            st.stop()

        # 3. 建立資料表
        clean_df = pd.DataFrame()
        clean_df['ID'] = df.iloc[:, id_idx]
        clean_df['IssueDate'] = pd.to_datetime(df.iloc[:, issue_date_idx], errors='coerce') if issue_date_idx else pd.Timestamp.min
        clean_df['ValuationDate'] = pd.to_datetime(df.iloc[:, final_date_idx], errors='coerce') if final_date_idx else pd.Timestamp.max
        
        clean_df['KO_Pct'] = df.iloc[:, ko_idx].apply(clean_percentage)
        clean_df['KI_Pct'] = df.iloc[:, ki_idx].apply(clean_percentage)
        clean_df['Strike_Pct'] = df.iloc[:, strike_idx].apply(clean_percentage) if strike_idx else 100.0
        
        clean_df['KO_Type'] = df.iloc[:, ko_type_idx] if ko_type_idx else ""
        clean_df['KI_Type'] = df.iloc[:, ki_type_idx] if ki_type_idx else "AKI"
        
        clean_df['Email'] = df.iloc[:, email_idx] if email_idx else ""
        clean_df['Name'] = df.iloc[:, name_idx] if name_idx else "客戶"
        
        for i in range(1, 6):
            if i == 1: tx_idx = t1_idx
            else:
                tx_idx, _ = find_col_index(cols, [f"標的{i}"])
                if tx_idx is None: tx_idx = t1_idx + (i-1)*2
            
            clean_df[f'T{i}_Code'] = df.iloc[:, tx_idx]
            clean_df[f'T{i}_Strike'] = df.iloc[:, tx_idx + 1]

        clean_df = clean_df.dropna(subset=['ID'])
        
        # 4. 抓取股價
        st.info("下載歷史資料回測中... ☕")
        all_tickers = []
        for i in range(1, 6):
            tickers = clean_df[f'T{i}_Code'].dropna().astype(str).unique().tolist()
            all_tickers.extend(tickers)
        all_tickers = [t.strip() for t in set(all_tickers) if t != 'nan' and str(t).strip() != '']
        
        if not all_tickers: st.stop()
            
        min_issue_date = clean_df['IssueDate'].min()
        if pd.isna(min_issue_date): min_issue_date = datetime.now() - timedelta(days=365)
        
        try:
            history_data = yf.download(all_tickers, start=min_issue_date)['Close']
        except:
            st.error("美股連線失敗")
            st.stop()

        # 5. 核心邏輯
        results = []
        today = pd.Timestamp.now()

        for index, row in clean_df.iterrows():
            ko_thresh_val = row['KO_Pct'] if pd.notna(row['KO_Pct']) else 100.0
            ki_thresh_val = row['KI_Pct'] if pd.notna(row['KI_Pct']) else 60.0
            strike_thresh_val = row['Strike_Pct'] if pd.notna(row['Strike_Pct']) else 100.0
            
            ko_thresh = ko_thresh_val / 100.0
            ki_thresh = ki_thresh_val / 100.0
            strike_thresh = strike_thresh_val / 100.0

            nc_months = parse_nc_months(row['KO_Type'])
            nc_end_date = row['IssueDate'] + relativedelta(months=nc_months)
            
            assets = []
            for i in range(1, 6):
                code = str(row[f'T{i}_Code']).strip()
                try: initial = float(row[f'T{i}_Strike'])
                except: initial = 0
                if code != 'nan' and code != '' and initial > 0:
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

            # --- 回測引擎 ---
            if len(all_tickers) == 1: product_history = history_data
            else: product_history = history_data[[a['code'] for a in assets]]
            
            sim_data = product_history[product_history.index >= row['IssueDate']]
            
            product_status = "Running"
            early_redemption_date = None
            is_aki = "AKI" in str(row['KI_Type']).upper()
            
            for date, prices in sim_data.iterrows():
                if product_status == "Early Redemption": break
                is_post_nc = date >= nc_end_date
                all_locked = True
                
                for asset in assets:
                    try:
                        if len(assets) == 1 and len(all_tickers) == 1: price = prices
                        else: price = prices[asset['code']]
                    except: continue 
                    
                    if pd.isna(price): continue
                    perf = price / asset['initial']
                    date_str = date.strftime('%Y/%m/%d')
                    
                    if is_aki and perf < ki_thresh:
                        if not asset['hit_ki']:
                            asset['hit_ki'] = True
                            asset['ki_record'] = f"@{price:.2f} ({date_str})"
                        
                    if not asset['locked_ko']:
                        if is_post_nc and perf >= ko_thresh:
                            asset['locked_ko'] = True 
                            asset['ko_record'] = f"@{price:.2f} ({date_str})"
                    
                    if not asset['locked_ko']: all_locked = False
                        
                if all_locked:
                    product_status = "Early Redemption"
                    early_redemption_date = date
            
            # --- 最終計算與整理 ---
            locked_list = []
            waiting_list = []
            hit_ki_list = []
            
            detail_lines = [] # 用來存合併的欄位資訊

            for asset in assets:
                try:
                    if len(all_tickers) == 1: curr = float(history_data.iloc[-1])
                    else: curr = float(history_data.iloc[-1][asset['code']])
                    asset['price'] = curr
                    asset['perf'] = curr / asset['initial']
                    if not is_aki and asset['perf'] < ki_thresh: 
                        asset['hit_ki'] = True
                        asset['ki_record'] = f"@{curr:.2f} (EKI)"
                except: pass
                
                # 分類
                if asset['locked_ko']: locked_list.append(asset['code'])
                else: waiting_list.append(asset['code'])
                
                if asset['hit_ki']: hit_ki_list.append(asset['code'])
                
                # 建立合併欄位的文字 (Code + Perf + Icon)
                p_pct = round(asset['perf']*100, 2)
                status_icon = "✅" if asset['locked_ko'] else "⚠️" if asset['hit_ki'] else ""
                
                # 格式：[AAPL] 105% ✅ (換行) KO @...
                line_str = f"[{asset['code']}] {p_pct}% {status_icon}"
                if asset['locked_ko']: line_str += f" (KO {asset['ko_record']})"
                if asset['hit_ki']: line_str += f" (KI {asset['ki_record']})"
                
                detail_lines.append(line_str)

            hit_any_ki = any(a['hit_ki'] for a in assets)
            all_above_strike_now = all(a['perf'] >= strike_thresh for a in assets)
            worst_asset = min(assets, key=lambda x: x['perf'])
            worst_perf = worst_asset['perf']
            
            # --- 狀態總結生成 (Smart Status) ---
            final_status = ""
            
            if today < row['IssueDate']:
                final_status = "⏳ 未發行"
            elif product_status == "Early Redemption":
                final_status = f"🎉 提前出場\n({early_redemption_date.strftime('%Y-%m-%d')})"
            elif pd.notna(row['ValuationDate']) and today >= row['ValuationDate']:
                if all_above_strike_now:
                     final_status = "💰 到期獲利\n(全數 > 執行價)"
                elif hit_any_ki:
                     final_status = f"😭 到期接股\n{worst_asset['code']} @ {round(worst_asset['strike_price'], 2)}"
                else:
                     final_status = "🛡️ 到期保本\n(未破KI)"
            else:
                # 存續期間 - 智慧摘要
                status_parts = []
                if today < nc_end_date:
                    status_parts.append(f"🔒 NC閉鎖中")
                else:
                    # 顯示誰鎖定了，誰還在等
                    wait_str = ",".join(waiting_list)
                    lock_str = ",".join(locked_list)
                    
                    if not waiting_list:
                        status_parts.append("👀 比價中")
                    else:
                        sub_msg = f"⏳等待: {wait_str}"
                        if locked_list:
                            sub_msg = f"✅已鎖: {lock_str}\n" + sub_msg
                        status_parts.append(f"👀 比價中\n{sub_msg}")
                
                if hit_any_ki:
                    ki_str = ",".join(hit_ki_list)
                    status_parts.append(f"⚠️ KI已破: {ki_str}")
                
                final_status = "\n".join(status_parts)

            # 準備輸出
            email_table = "【標的詳細狀態】\n"
            email_table += f"{'代碼':<6} | {'KO紀錄':<18} | {'現價':<8} | {'KI紀錄':<18}\n"
            email_table += "-"*60 + "\n"
            for asset in assets:
                 ko_info = asset['ko_record'] if asset['locked_ko'] else ".."
                 ki_info = asset['ki_record'] if asset['hit_ki'] else ""
                 email_table += f"{asset['code']:<6} | {ko_info:<18} | {round(asset['price'], 2):<8} | {ki_info:<18}\n"

            row_res = {
                "債券代號": row['ID'],
                "收件人": row['Name'],
                "Email": str(row['Email']).strip(),
                "發行日": row['IssueDate'].strftime('%Y-%m-%d'),
                "狀態": final_status,
                "最差表現": f"{round(worst_perf*100, 2)}%",
                "KO設定": f"{ko_thresh_val}%",
                "KI設定": f"{ki_thresh_val}%",
                "執行價": f"{strike_thresh_val}%",
                # 將多個標的資訊合併成一個欄位
                "標的明細 (代碼/表現/紀錄)": "\n".join(detail_lines),
                
                "msg_subject": f"【ELN通知】{row['ID']} 狀態更新",
                "msg_body": (
                    f"Hi {row['Name']}：\n\n"
                    f"商品 {row['ID']} 最新報告：\n"
                    f"📊 狀態：\n{final_status}\n\n"
                    f"⚡ 設定：KO {ko_thresh_val}% / KI {ki_thresh_val}% ({row['KI_Type']})\n"
                    f"📉 執行價格(Strike)：{strike_thresh_val}%\n\n"
                    f"{email_table}\n"
                    f"--------------------------------\n"
                    f"系統自動發送"
                )
            }
            results.append(row_res)

        # 6. 顯示
        final_df = pd.DataFrame(results)
        
        st.subheader("📋 專業監控列表")
        
        def color_status(val):
            if "提前" in str(val) or "獲利" in str(val): return 'background-color: #d4edda; color: green'
            if "接股" in str(val) or "KI" in str(val): return 'background-color: #f8d7da; color: red'
            if "NC" in str(val) or "未發行" in str(val): return 'background-color: #fff3cd; color: #856404'
            return ''

        # 這裡選擇要顯示的欄位，不再顯示 T1_狀態... T5_狀態
        display_cols = ['債券代號', '狀態', '最差表現', '標的明細 (代碼/表現/紀錄)', 'KO設定', 'KI設定', '執行價', '發行日']
        
        column_config = {
            "狀態": st.column_config.TextColumn("目前狀態摘要", width="large", help="顯示目前鎖定與等待進度"),
            "債券代號": st.column_config.TextColumn("代號", width="small"),
            "KO設定": st.column_config.TextColumn("KO", width="small"),
            "KI設定": st.column_config.TextColumn("KI", width="small"),
            "執行價": st.column_config.TextColumn("Strike", width="small"),
            "最差表現": st.column_config.TextColumn("Worst Of", width="small"),
            "標的明細 (代碼/表現/紀錄)": st.column_config.TextColumn("標的明細 (Code / Perf / History)", width="large"),
        }

        st.dataframe(
            final_df[display_cols].style.applymap(color_status, subset=['狀態']), 
            use_container_width=True,
            column_config=column_config,
            height=600, # 表格加高，因為現在內容有換行
            hide_index=True
        )
        
        st.markdown("### 📢 發信操作")
        edited_df = st.data_editor(final_df[['債券代號', '收件人', 'Email', '狀態']], key='editor')
        
        for idx, row in final_df.iterrows():
            if any(x in row['狀態'] for x in ["提前", "到期", "已破", "獲利", "接股"]):
                email = edited_df.iloc[idx]['Email']
                if st.button(f"📧 通知 {row['債券代號']}", key=f"btn_{idx}"):
                    if sender_email:
                        send_email(sender_email, sender_password, email, row['msg_subject'], row['msg_body'])
                    else:
                        st.error("請填寫寄件人資訊")

    except Exception as e:
        st.error(f"發生錯誤：{e}")
else:
    st.info("👆 請上傳 Excel")
