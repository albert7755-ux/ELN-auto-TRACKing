import streamlit as st
import pandas as pd
import yfinance as yf
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# --- 設定網頁 ---
st.set_page_config(page_title="ELN 自動戰情室 (詳細版)", layout="wide")

# --- 側邊欄：設定 Email 寄件資訊 ---
with st.sidebar:
    st.header("📧 Email 設定中心")
    sender_email = st.text_input("寄件人 Gmail", placeholder="example@gmail.com")
    sender_password = st.text_input("應用程式密碼", type="password", placeholder="16位數密碼")
    st.info("💡 修正更新：現在會顯示詳細價格數據，並判斷發行日。")

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

# --- 智慧搜尋欄位函數 ---
def find_col_index(columns, keywords):
    for idx, col_name in enumerate(columns):
        col_str = str(col_name).strip().lower()
        if any(k in col_str for k in keywords):
            return idx
    return None

# --- 主畫面 ---
st.title("📊 ELN 結構型商品 - 自動監控戰情室")
st.markdown("### 🔍 詳細數據版 (含進場價、現價、KO/KI 資訊)")

uploaded_file = st.file_uploader("請上傳 Excel 檔案 (工作表1格式)", type=['xlsx'])

if uploaded_file is not None:
    try:
        # 1. 讀取資料
        try:
            df = pd.read_excel(uploaded_file, sheet_name=0, header=0, engine='openpyxl')
        except:
            df = pd.read_excel(uploaded_file, sheet_name=0, header=0)

        # 🧹 資料清洗：移除「進場價」那一行中文標題
        # 檢查第一列是否包含 "進場價" 這種字眼，有的話就刪掉
        if df.iloc[0].astype(str).str.contains("進場價").any():
            df = df.iloc[1:].reset_index(drop=True)

        cols = df.columns.tolist()
        
        # --- 2. 智慧定位欄位 ---
        id_idx = find_col_index(cols, ["債券", "代號", "id"]) or 0
        ko_idx = find_col_index(cols, ["ko", "價格"]) or find_col_index(cols, ["ko", "%"])
        ki_idx = find_col_index(cols, ["ki", "價格"]) or find_col_index(cols, ["ki", "%"])
        t1_idx = find_col_index(cols, ["標的1"])
        
        # 尋找日期欄位 (發行日)
        date_idx = find_col_index(cols, ["發行日", "交易日", "date"])
        
        # Email 與 姓名
        email_idx = find_col_index(cols, ["email", "信箱"])
        name_idx = find_col_index(cols, ["理專", "姓名", "客戶"])

        if t1_idx is None or ko_idx is None:
            st.error("❌ 無法辨識關鍵欄位，請確認 Excel 標題包含「債券代號」、「標的1」、「KO」。")
            st.stop()

        # --- 3. 建立乾淨的資料表 ---
        clean_df = pd.DataFrame()
        clean_df['ID'] = df.iloc[:, id_idx]
        
        # 日期處理
        if date_idx:
            clean_df['StartDate'] = pd.to_datetime(df.iloc[:, date_idx], errors='coerce')
        else:
            clean_df['StartDate'] = pd.Timestamp.min # 沒日期就預設很早

        # 其他欄位
        clean_df['Email'] = df.iloc[:, email_idx] if email_idx else ""
        clean_df['Name'] = df.iloc[:, name_idx] if name_idx else "客戶"
        clean_df['KO_Pct'] = df.iloc[:, ko_idx]
        clean_df['KI_Pct'] = df.iloc[:, ki_idx] if ki_idx else 60.0
        
        # 抓取標的 1~5 (代碼 + 進場價)
        # 假設結構是：[標的1代碼] [標的1進場價] [標的2代碼] ...
        for i in range(1, 6):
            if i == 1:
                tx_idx = t1_idx
            else:
                # 嘗試搜尋
                found = find_col_index(cols, [f"標的{i}"])
                tx_idx = found if found else t1_idx + (i-1)*2
            
            clean_df[f'T{i}_Code'] = df.iloc[:, tx_idx]
            clean_df[f'T{i}_Strike'] = df.iloc[:, tx_idx + 1]

        clean_df = clean_df.dropna(subset=['ID'])
        
        # --- 4. 抓股價 ---
        st.info("連線美股報價中... ☕")
        all_tickers = []
        for i in range(1, 6):
            tickers = clean_df[f'T{i}_Code'].dropna().astype(str).unique().tolist()
            all_tickers.extend(tickers)
        all_tickers = [t.strip() for t in set(all_tickers) if t != 'nan' and str(t).strip() != '']
        
        if not all_tickers:
            st.error("找不到股票代碼")
            st.stop()
            
        market_data = yf.download(all_tickers, period="1d")['Close']
        if not market_data.empty:
            latest_prices = market_data.iloc[-1]
        else:
            st.error("無法抓取股價")
            st.stop()

        # --- 5. 核心計算 (含詳細資訊) ---
        results = []
        today = pd.Timestamp.now()

        for index, row in clean_df.iterrows():
            row_output = {
                "債券代號": row['ID'],
                "收件人": row['Name'],
                "Email": str(row['Email']).strip() if pd.notna(row['Email']) else "",
                "發行日": row['StartDate'].strftime('%Y-%m-%d') if pd.notna(row['StartDate']) else "N/A",
                "KO價": f"{row['KO_Pct']}%",
                "KI價": f"{row['KI_Pct']}%",
                "狀態": "觀察中",
                "最差表現": 0.0,
                "msg_body": ""
            }
            
            # 0. 尚未比價判斷
            if pd.notna(row['StartDate']) and today < row['StartDate']:
                row_output["狀態"] = "⏳ 尚未比價 (未發行)"
                results.append(row_output)
                continue # 跳過後續計算
            
            try:
                ko_threshold = float(row['KO_Pct']) / 100
                ki_threshold = float(row['KI_Pct']) / 100
            except:
                ko_threshold = 1.0; ki_threshold = 0.6
                
            perfs = []
            is_all_ko = True
            hit_ki = False
            
            # 用來做 Email 的表格文字
            email_table = "【詳細標的資訊】\n"
            email_table += f"{'代碼':<6} | {'現價':<8} | {'進場價':<8} | {'表現(%)':<8}\n"
            email_table += "-"*45 + "\n"
            
            for i in range(1, 6):
                code = str(row[f'T{i}_Code']).strip()
                try:
                    initial = float(row[f'T{i}_Strike'])
                except:
                    initial = 0
                
                if code == 'nan' or code == '' or initial == 0:
                    continue
                
                try:
                    if len(all_tickers) == 1:
                        curr = float(latest_prices)
                    else:
                        curr = float(latest_prices[code])
                    
                    p = curr / initial
                    perfs.append(p)
                    
                    if p < ko_threshold: is_all_ko = False
                    if p < ki_threshold: hit_ki = True
                    
                    # 存入結果表 (給網頁顯示用)
                    p_pct = round(p * 100, 2)
                    row_output[f"T{i}_代碼"] = code
                    row_output[f"T{i}_進場"] = initial
                    row_output[f"T{i}_現價"] = round(curr, 2)
                    row_output[f"T{i}_表現"] = f"{p_pct}%"
                    
                    # 存入 Email 文字
                    email_table += f"{code:<6} | {round(curr, 2):<8} | {initial:<8} | {p_pct:<8}\n"
                    
                except:
                    row_output[f"T{i}_表現"] = "Error"
                    is_all_ko = False

            if perfs:
                worst = min(perfs)
                row_output["最差表現"] = f"{round(worst*100, 2)}%"
                
                status_msg = "👀 觀察中"
                if is_all_ko: status_msg = "🎉 獲利出場 (KO)"
                elif hit_ki: status_msg = "⚠️ 下檔保護失效 (HIT)"
                
                row_output["狀態"] = status_msg
                
                # 準備信件內容
                row_output["msg_subject"] = f"【ELN通知】{row['ID']} 狀態：{status_msg}"
                row_output["msg_body"] = (
                    f"Hi {row['Name']}：\n\n"
                    f"您關注的商品 {row['ID']} 最新監控報告：\n"
                    f"📅 發行日：{row_output['發行日']}\n"
                    f"📊 目前狀態：{status_msg}\n"
                    f"📉 最差表現：{round(worst*100, 2)}%\n"
                    f"🚩 KO門檻：{row['KO_Pct']}%\n"
                    f"🛡️ KI門檻：{row['KI_Pct']}%\n\n"
                    f"{email_table}\n"
                    f"--------------------------------\n"
                    f"(本郵件由自動化系統發送)"
                )

            results.append(row_output)

        # --- 6. 顯示結果 ---
        final_df = pd.DataFrame(results)
        
        st.subheader("📋 詳細監控列表")
        st.caption("以下列表已展開所有標的資訊")
        
        # 整理顯示欄位順序
        base_cols = ['債券代號', '收件人', '狀態', '最差表現', 'KO價', 'KI價', '發行日', 'Email']
        detail_cols = [c for c in final_df.columns if c.startswith('T') and c not in base_cols]
        # 排序 detail_cols
        detail_cols.sort()
        
        st.dataframe(final_df[base_cols + detail_cols], use_container_width=True)
        
        # --- 發信區 ---
        st.markdown("### 📢 一鍵發信")
        
        edited_df = st.data_editor(
            final_df[['債券代號', '收件人', 'Email', '狀態']],
            column_config={"Email": st.column_config.TextColumn("Email")},
            use_container_width=True,
            num_rows="fixed",
            key="email_editor"
        )
        
        for idx, row in final_df.iterrows():
            if "KO" in row['狀態'] or "HIT" in row['狀態']:
                current_email = edited_df.iloc[idx]['Email']
                
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.text(f"通知 {row['收件人']} ({current_email}) - {row['狀態']}")
                with col2:
                    if sender_email and current_email:
                        if st.button(f"📧 發信", key=f"mail_{idx}"):
                            send_email(sender_email, sender_password, current_email, row['msg_subject'], row['msg_body'])
                    else:
                        st.button("🚫 缺資料", disabled=True, key=f"dis_{idx}")

    except Exception as e:
        st.error(f"發生錯誤：{e}")
else:
    st.info("👆 請上傳 Excel")
