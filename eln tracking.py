import streamlit as st
import pandas as pd
import yfinance as yf
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- 設定網頁 ---
st.set_page_config(page_title="ELN 自動戰情室 (Email版)", layout="wide")

# --- 側邊欄：設定 Email 寄件資訊 ---
with st.sidebar:
    st.header("📧 Email 設定中心")
    st.markdown("請輸入您的 Gmail 寄件資訊")
    
    sender_email = st.text_input("寄件人 Gmail", placeholder="example@gmail.com")
    sender_password = st.text_input("應用程式密碼", type="password", placeholder="16位數密碼", help="請至 Google 帳戶 > 安全性 > 應用程式密碼 申請")
    
    st.info("💡 程式會自動偵測 Excel 中的「Email」欄位來發信。")

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
st.caption("🚀 支援 Excel 自動匯入 Email 名單 (請在 Excel 新增 'Email' 欄位)")

uploaded_file = st.file_uploader("請上傳 Excel 檔案 (工作表1格式)", type=['xlsx'])

if uploaded_file is not None:
    try:
        # 1. 讀取資料 (使用 openpyxl 引擎)
        try:
            df = pd.read_excel(uploaded_file, sheet_name=0, header=0, engine='openpyxl')
        except:
            df = pd.read_excel(uploaded_file, sheet_name=0, header=0)

        cols = df.columns.tolist()
        
        # --- 2. 智慧定位欄位 ---
        id_idx = find_col_index(cols, ["債券", "代號", "id"]) or 0
        ko_idx = find_col_index(cols, ["ko", "%"]) or find_col_index(cols, ["ko", "價格"])
        ki_idx = find_col_index(cols, ["ki", "%"]) or find_col_index(cols, ["ki", "價格"])
        t1_idx = find_col_index(cols, ["標的1"])
        
        # 尋找 Email 欄位 (支援多種寫法)
        email_idx = find_col_index(cols, ["email", "信箱", "郵件", "e-mail"])
        # 尋找 姓名/理專 欄位 (選填)
        name_idx = find_col_index(cols, ["理專", "姓名", "客戶", "name"])

        if t1_idx is None or ko_idx is None:
            st.error("❌ 無法辨識關鍵欄位，請確認 Excel 標題包含「債券代號」、「標的1」、「KO」。")
            st.stop()

        # --- 3. 建立資料表 ---
        clean_df = pd.DataFrame()
        clean_df['ID'] = df.iloc[:, id_idx]
        
        # 處理 Email
        if email_idx is not None:
            clean_df['Email'] = df.iloc[:, email_idx]
        else:
            clean_df['Email'] = "" # 沒找到欄位就留白
            
        # 處理 姓名
        if name_idx is not None:
            clean_df['Name'] = df.iloc[:, name_idx]
        else:
            clean_df['Name'] = "客戶"

        # 抓取數值
        clean_df['KO_Pct'] = df.iloc[:, ko_idx]
        clean_df['KI_Pct'] = df.iloc[:, ki_idx] if ki_idx else 60.0
        
        # 抓取標的 1~5
        clean_df['T1_Code'] = df.iloc[:, t1_idx]
        clean_df['T1_Strike'] = df.iloc[:, t1_idx + 1] # 進場價通常在代碼右邊
        
        # 簡易迴圈抓 T2~T5 (智慧判斷)
        for i in range(2, 6):
            tx_idx = find_col_index(cols, [f"標的{i}"])
            if tx_idx:
                clean_df[f'T{i}_Code'] = df.iloc[:, tx_idx]
                clean_df[f'T{i}_Strike'] = df.iloc[:, tx_idx + 1]
            else:
                # 找不到就用推算的 (假設每2格一組)
                offset = (i-1) * 2
                clean_df[f'T{i}_Code'] = df.iloc[:, t1_idx + offset]
                clean_df[f'T{i}_Strike'] = df.iloc[:, t1_idx + offset + 1]

        clean_df = clean_df.dropna(subset=['ID'])
        
        # --- 4. 抓股價 ---
        st.info("連線美股報價中... ☕")
        all_tickers = []
        for i in range(1, 6):
            tickers = clean_df[f'T{i}_Code'].dropna().astype(str).unique().tolist()
            all_tickers.extend(tickers)
        all_tickers = [t.strip() for t in set(all_tickers) if t != 'nan' and str(t).strip() != '']
        
        if not all_tickers:
            st.error("Excel 中找不到任何股票代碼")
            st.stop()
            
        market_data = yf.download(all_tickers, period="1d")['Close']
        if not market_data.empty:
            latest_prices = market_data.iloc[-1]
        else:
            st.error("無法抓取股價")
            st.stop()

        # --- 5. 計算結果 ---
        results = []
        for index, row in clean_df.iterrows():
            row_output = {
                "債券代號": row['ID'],
                "收件人": row['Name'],
                "Email": str(row['Email']).strip() if pd.notna(row['Email']) else "",
                "狀態": "觀察中",
                "最差表現": 0.0,
                "msg_subject": "",
                "msg_body": ""
            }
            
            try:
                ko_threshold = float(row['KO_Pct']) / 100
                ki_threshold = float(row['KI_Pct']) / 100
            except:
                ko_threshold = 1.0
                ki_threshold = 0.6
                
            perfs = []
            is_all_ko = True
            hit_ki = False
            details_text = ""
            
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
                    
                    row_output[f"標的{i}"] = code
                    row_output[f"現價{i}"] = round(curr, 2)
                    row_output[f"表現{i}"] = f"{round(p*100, 2)}%"
                    details_text += f"- {code}: 現價 {round(curr, 2)} / 進場 {initial} ({round(p*100, 2)}%)\n"
                except:
                    pass

            if perfs:
                worst = min(perfs)
                row_output["最差表現"] = f"{round(worst*100, 2)}%"
                
                status_msg = "👀 觀察中"
                if is_all_ko: status_msg = "🎉 獲利出場 (KO)"
                elif hit_ki: status_msg = "⚠️ 下檔保護失效 (HIT)"
                
                row_output["狀態"] = status_msg
                
                # 準備信件內容
                row_output["msg_subject"] = f"【ELN通知】{row['ID']} 最新狀態：{status_msg}"
                row_output["msg_body"] = (
                    f"Hi {row['Name']}：\n\n"
                    f"您關注的商品 {row['ID']} 今日狀態更新：\n"
                    f"狀態：{status_msg}\n"
                    f"最差表現：{round(worst*100, 2)}%\n"
                    f"--------------------------------\n"
                    f"{details_text}\n"
                    f"(本郵件由系統自動發送)"
                )

            results.append(row_output)

        # --- 6. 顯示結果 ---
        final_df = pd.DataFrame(results)
        
        st.subheader("📋 監控與發信列表")
        st.caption("程式會自動抓取 Excel 中的 Email，您也可以在下方直接修改後發送。")
        
        # 讓使用者可以臨時修改 Email (使用 Data Editor)
        edited_df = st.data_editor(
            final_df[['債券代號', '收件人', 'Email', '狀態', '最差表現']],
            column_config={
                "Email": st.column_config.TextColumn("Email (可編輯)", help="填入收信者的 Email"),
            },
            use_container_width=True,
            num_rows="fixed"
        )
        
        st.markdown("### 📢 一鍵發信")
        
        # 找出建議通知的項目
        for idx, row in final_df.iterrows():
            if "KO" in row['狀態'] or "HIT" in row['狀態']:
                # 取得在上方表格可能被修改過的 Email
                current_email = edited_df.iloc[idx]['Email']
                current_name = edited_df.iloc[idx]['收件人']
                
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.text(f"通知 {current_name} ({current_email}) - {row['狀態']}")
                with col2:
                    if sender_email and sender_password and current_email:
                        if st.button(f"📧 發信", key=f"mail_{idx}"):
                            send_email(
                                sender_email, 
                                sender_password, 
                                current_email, 
                                row['msg_subject'], 
                                row['msg_body']
                            )
                    else:
                        st.button(f"🚫 資料不全", key=f"dis_{idx}", disabled=True, help="請確認側邊欄已填寫寄件資訊，且該筆資料有 Email")

    except Exception as e:
        st.error(f"發生錯誤：{e}")
else:
    st.info("👆 請上傳 Excel")
