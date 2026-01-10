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
    st.info("💡 **邏輯更新：**\n1. 到期接股時，顯示標的與執行價\n2. 支援獨立記憶 KO\n3. 每日路徑回測 AKI")

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
        return 1 # 預設 NC1M
    match = re.search(r'NC(\d+)', str(ko_type_str), re.IGNORECASE)
    if match:
        return int(match.group(1))
    return 1 

# --- 函數：尋找欄位 ---
def find_col_index(columns, keywords):
    for idx, col_name in enumerate(columns):
        col_str = str(col_name).strip().lower()
        if any(k in col_str for k in keywords):
            return idx
    return None

# --- 主畫面 ---
st.title("📊 ELN 結構型商品 - 專業監控戰情室")
st.markdown("### 🚀 支援到期接股明細 (標的/執行價)、獨立 KO 鎖定")

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
        
        # 2. 定位欄位
        id_idx = find_col_index(cols, ["債券", "代號"]) or 0
        ko_idx = find_col_index(cols, ["ko", "價格"]) or find_col_index(cols, ["ko", "%"])
        ko_type_idx = find_col_index(cols, ["ko", "類型", "type"])
        ki_idx = find_col_index(cols, ["ki", "價格"]) or find_col_index(cols, ["ki", "%"])
        ki_type_idx = find_col_index(cols, ["ki", "類型", "type"])
        strike_idx = find_col_index(cols, ["執行", "strike"]) 
        t1_idx = find_col_index(cols, ["標的1"])
        issue_date_idx = find_col_index(cols, ["發行日"])
        final_date_idx = find_col_index(cols, ["最終", "評價", "final"])
        email_idx = find_col_index(cols, ["email", "信箱"])
        name_idx = find_col_index(cols, ["理專", "姓名", "客戶"])

        if t1_idx is None or ko_idx is None:
            st.error("❌ 欄位辨識失敗")
            st.stop()

        # 3. 建立資料表
        clean_df = pd.DataFrame()
        clean_df['ID'] = df.iloc[:, id_idx]
        clean_df['IssueDate'] = pd.to_datetime(df.iloc[:, issue_date_idx], errors='coerce') if issue_date_idx else pd.Timestamp.min
        clean_df['ValuationDate'] = pd.to_datetime(df.iloc[:, final_date_idx], errors='coerce') if final_date_idx else pd.Timestamp.max
        clean_df['KO_Pct'] = pd.to_numeric(df.iloc[:, ko_idx], errors='coerce')
        clean_df['KO_Type'] = df.iloc[:, ko_type_idx] if ko_type_idx else ""
        clean_df['KI_Pct'] = pd.to_numeric(df.iloc[:, ki_idx], errors='coerce')
        clean_df['KI_Type'] = df.iloc[:, ki_type_idx] if ki_type_idx else "AKI"
        clean_df['Strike_Pct'] = pd.to_numeric(df.iloc[:, strike_idx], errors='coerce') if strike_idx else 100.0
        clean_df['Email'] = df.iloc[:, email_idx] if email_idx else ""
        clean_df['Name'] = df.iloc[:, name_idx] if name_idx else "客戶"
        
        for i in range(1, 6):
            if i == 1: tx_idx = t1_idx
            else:
                found = find_col_index(cols, [f"標的{i}"])
                tx_idx = found if found else t1_idx + (i-1)*2
            clean_df[f'T{i}_Code'] = df.iloc[:, tx_idx]
            clean_df[f'T{i}_Strike'] = df.iloc[:, tx_idx + 1]

        clean_df = clean_df.dropna(subset=['ID'])
        
        # 4. 抓取股價
        st.info("下載歷史資料進行路徑回測... ☕")
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

        # 5. 核心邏輯 (每日回測模擬)
        results = []
        today = pd.Timestamp.now()

        for index, row in clean_df.iterrows():
            try:
                ko_thresh = float(row['KO_Pct']) / 100
                ki_thresh = float(row['KI_Pct']) / 100
                strike_thresh = float(row['Strike_Pct']) / 100
            except:
                ko_thresh = 1.0; ki_thresh = 0.6; strike_thresh = 1.0

            # 參數
            nc_months = parse_nc_months(row['KO_Type'])
            nc_end_date = row['IssueDate'] + relativedelta(months=nc_months)
            
            # 初始化資產狀態
            assets = []
            for i in range(1, 6):
                code = str(row[f'T{i}_Code']).strip()
                try: initial = float(row[f'T{i}_Strike'])
                except: initial = 0
                if code != 'nan' and code != '' and initial > 0:
                    assets.append({
                        'code': code, 
                        'initial': initial,
                        'strike_price': initial * strike_thresh, # 計算該股的執行價數值
                        'locked_ko': False, 
                        'hit_ki': False,
                        'perf': 0.0, 
                        'price': 0.0
                    })
            
            if not assets: continue

            # --- 模擬回測引擎 ---
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
                    
                    # AKI 檢查
                    if is_aki and perf < ki_thresh:
                        asset['hit_ki'] = True
                        
                    # 獨立 KO 檢查
                    if not asset['locked_ko']:
                        if is_post_nc and perf >= ko_thresh:
                            asset['locked_ko'] = True 
                    
                    if not asset['locked_ko']: all_locked = False
                        
                if all_locked:
                    product_status = "Early Redemption"
                    early_redemption_date = date
            
            # --- 取得最終狀態 ---
            for asset in assets:
                try:
                    if len(all_tickers) == 1: curr = float(history_data.iloc[-1])
                    else: curr = float(history_data.iloc[-1][asset['code']])
                    asset['price'] = curr
                    asset['perf'] = curr / asset['initial']
                    # EKI 檢查
                    if not is_aki and asset['perf'] < ki_thresh:
                        asset['hit_ki'] = True
                except: pass

            hit_any_ki = any(a['hit_ki'] for a in assets)
            all_above_strike_now = all(a['perf'] >= strike_thresh for a in assets)
            
            # 找出表現最差的標的 (為了接股準備)
            worst_asset = min(assets, key=lambda x: x['perf'])
            worst_perf = worst_asset['perf']
            
            final_status = ""
            
            # 1. 尚未發行
            if today < row['IssueDate']:
                final_status = "⏳ 未發行"
            
            # 2. 已提前出場 (KO)
            elif product_status == "Early Redemption":
                final_status = f"🎉 提前出場 (於 {early_redemption_date.strftime('%Y-%m-%d')})"
            
            # 3. 最終評價日 (到期結算)
            elif pd.notna(row['ValuationDate']) and today >= row['ValuationDate']:
                if all_above_strike_now:
                     final_status = "💰 到期獲利 (全數 > 執行價)"
                elif hit_any_ki:
                     # 發生接股：顯示哪一支、執行價多少
                     final_status = f"😭 到期接股: {worst_asset['code']} (執行價 {round(worst_asset['strike_price'], 2)})"
                else:
                     final_status = "🛡️ 到期保本 (未破KI)"
            
            # 4. 存續期間 (觀察中)
            else:
                locked_count = sum(1 for a in assets if a['locked_ko'])
                status_parts = []
                if today < nc_end_date:
                    status_parts.append(f"🔒 NC閉鎖中")
                else:
                    status_parts.append(f"👀 比價中 (KO: {locked_count}/{len(assets)})")
                
                if hit_any_ki:
                    status_parts.append("⚠️ AKI已破")
                
                final_status = " ".join(status_parts)

            # 準備輸出
            email_table = "【標的詳細狀態】\n"
            email_table += f"{'代碼':<6} | {'KO鎖定':<6} | {'現價':<8} | {'進場價':<8} | {'表現(%)':<8}\n"
            email_table += "-"*55 + "\n"
            
            detail_cols = {}
            for i, asset in enumerate(assets):
                lock_icon = "✅" if asset['locked_ko'] else ".."
                ki_icon = "⚠️" if asset['hit_ki'] else ""
                p_pct = round(asset['perf']*100, 2)
                email_table += f"{asset['code']:<6} | {lock_icon:<6} | {round(asset['price'], 2):<8} | {asset['initial']:<8} | {p_pct:<8} {ki_icon}\n"
                detail_cols[f"T{i+1}_表現"] = f"{p_pct}% {lock_icon}"

            row_res = {
                "債券代號": row['ID'],
                "收件人": row['Name'],
                "Email": str(row['Email']).strip(),
                "發行日": row['IssueDate'].strftime('%Y-%m-%d'),
                "NC解鎖日": nc_end_date.strftime('%Y-%m-%d'),
                "狀態": final_status,
                "最差表現": f"{round(worst_perf*100, 2)}%",
                "msg_subject": f"【ELN通知】{row['ID']} 狀態：{final_status}",
                "msg_body": (
                    f"Hi {row['Name']}：\n\n"
                    f"商品 {row['ID']} 最新報告：\n"
                    f"📊 狀態：{final_status}\n"
                    f"📅 評價日：{row['ValuationDate'].strftime('%Y-%m-%d')}\n"
                    f"⚡ 條件：KO {row['KO_Pct']}% (獨立) / KI {row['KI_Pct']}% ({row['KI_Type']})\n"
                    f"📉 執行價格(Strike)：{row['Strike_Pct']}%\n\n"
                    f"{email_table}\n"
                    f"(✅=已KO鎖定, ⚠️=曾破KI)\n"
                    f"--------------------------------\n"
                    f"系統自動發送"
                )
            }
            row_res.update(detail_cols)
            results.append(row_res)

        # 6. 顯示
        final_df = pd.DataFrame(results)
        
        st.subheader("📋 專業監控列表 (含接股資訊)")
        
        def color_status(val):
            if "提前" in str(val) or "獲利" in str(val): return 'background-color: #d4edda; color: green'
            if "接股" in str(val) or "AKI" in str(val): return 'background-color: #f8d7da; color: red'
            if "NC" in str(val): return 'background-color: #fff3cd; color: #856404'
            return ''

        display_cols = ['債券代號', '狀態', '最差表現', 'NC解鎖日', '發行日'] + \
                       [c for c in final_df.columns if '表現' in c and '最差' not in c]
        
        st.dataframe(final_df[display_cols].style.applymap(color_status, subset=['狀態']), use_container_width=True)
        
        st.markdown("### 📢 發信操作")
        edited_df = st.data_editor(final_df[['債券代號', '收件人', 'Email', '狀態']], key='editor')
        
        for idx, row in final_df.iterrows():
            if any(x in row['狀態'] for x in ["提前", "到期", "AKI", "獲利", "接股"]):
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
