import streamlit as st
import pandas as pd
import yfinance as yf
import requests

# --- 設定網頁 ---
st.set_page_config(page_title="ELN 自動戰情室 (Line版)", layout="wide")

# --- 側邊欄：設定 Line Token ---
with st.sidebar:
    st.header("💬 Line 通知設定")
    st.markdown("請輸入您的 Line Notify 權杖")
    
    # 讓使用者輸入 Token (密碼形式)
    line_token = st.text_input("Line Token", type="password", placeholder="貼上剛剛申請的那串亂碼...")
    
    st.markdown("---")
    st.info("💡 **小撇步：**\n1. 去 [Line Notify](https://notify-bot.line.me/) 申請權杖\n2. 若要發到群組，記得邀請 'Line Notify' 機器人入群")

# --- 函數：發送 Line 通知 ---
def send_line_notify(token, message):
    if not token:
        st.warning("⚠️ 請先在左側輸入 Line Token")
        return False
    
    url = "https://notify-api.line.me/api/notify"
    headers = {"Authorization": "Bearer " + token}
    data = {"message": message}
    
    try:
        response = requests.post(url, headers=headers, data=data)
        if response.status_code == 200:
            st.toast("✅ Line 通知已發送！", icon="🚀")
            return True
        else:
            st.error(f"❌ 發送失敗，錯誤碼：{response.status_code}")
            return False
    except Exception as e:
        st.error(f"連線錯誤：{e}")
        return False

# --- 主畫面 ---
st.title("📊 ELN 結構型商品 - 自動監控戰情室")
st.markdown("### 💬 Line 通知專用版")

uploaded_file = st.file_uploader("請上傳 Excel 檔案 (工作表1)", type=['xlsx'])

if uploaded_file is not None:
    try:
        # 1. 讀取資料 (跳過第一列標題)
    
df = pd.read_excel(uploaded_file, sheet_name=0, header=1, engine='openpyxl')

        # 2. 建立乾淨的 DataFrame (對應你的工作表1格式)
        clean_df = pd.DataFrame()
        clean_df['ID'] = df.iloc[:, 0]  # 債券代號
        
        # 抓取 5 檔標的
        clean_df['T1_Code'] = df.iloc[:, 7]
        clean_df['T1_Strike'] = df.iloc[:, 8]
        clean_df['T2_Code'] = df.iloc[:, 9]
        clean_df['T2_Strike'] = df.iloc[:, 10]
        clean_df['T3_Code'] = df.iloc[:, 11]
        clean_df['T3_Strike'] = df.iloc[:, 12]
        clean_df['T4_Code'] = df.iloc[:, 13]
        clean_df['T4_Strike'] = df.iloc[:, 14]
        clean_df['T5_Code'] = df.iloc[:, 15]
        clean_df['T5_Strike'] = df.iloc[:, 16]
        
        clean_df['KO_Pct'] = df.iloc[:, 20]
        clean_df['KI_Pct'] = df.iloc[:, 22]
        
        clean_df = clean_df.dropna(subset=['ID'])
        
        # 3. 抓取美股現價
        st.info("連線美股報價中... 請稍候 ☕")
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

        # 4. 核心計算
        results = []
        for index, row in clean_df.iterrows():
            row_output = {
                "債券代號": row['ID'],
                "狀態": "觀察中",
                "最差表現": 0.0,
                "msg_content": ""
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
            details_text = "" # 用來組裝 Line 訊息
            
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
                    
                    icon = "✅" if p >= ko_threshold else "⚠️" if p < ki_threshold else ""
                    if p < ko_threshold: is_all_ko = False
                    if p < ki_threshold: hit_ki = True
                    
                    row_output[f"標的{i}"] = code
                    row_output[f"現價{i}"] = round(curr, 2)
                    row_output[f"表現{i}"] = f"{round(p*100, 2)}% {icon}"
                    
                    # Line 訊息要簡潔
                    details_text += f"{code}: {round(p*100, 1)}%\n"
                    
                except:
                    row_output[f"表現{i}"] = "Error"
                    is_all_ko = False

            if perfs:
                worst = min(perfs)
                row_output["最差表現"] = f"{round(worst*100, 2)}%"
                
                status_msg = "👀 觀察中"
                if is_all_ko: status_msg = "🎉 獲利出場 (KO)"
                elif hit_ki: status_msg = "⚠️ 下檔保護失效 (HIT)"
                
                row_output["狀態"] = status_msg
                
                # 組裝給 Line 的文字 (換行符號是 \n)
                row_output["msg_content"] = (
                    f"\n【ELN快訊】{row['ID']}\n"
                    f"狀態：{status_msg}\n"
                    f"最差表現：{round(worst*100, 2)}%\n"
                    f"----------------\n"
                    f"{details_text}"
                )

            results.append(row_output)

        # 5. 顯示結果
        final_df = pd.DataFrame(results)
        
        st.subheader("📋 監控列表")
        st.caption("勾選您想通知的商品，按下按鈕即可發送到 Line")

        # 使用 Streamlit 的表格呈現
        st.dataframe(
            final_df[['債券代號', '狀態', '最差表現'] + [c for c in final_df.columns if '標的' in c or '表現' in c]], 
            use_container_width=True
        )
        
        st.markdown("### 📢 發送通知區")
        
        # 只列出有 KO 或 HIT 的商品建議發送
        for idx, row in final_df.iterrows():
            if "KO" in row['狀態'] or "HIT" in row['狀態']:
                
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.text(f"建議發送：{row['債券代號']} - {row['狀態']}")
                with col2:
                    if st.button(f"💬 發 Line", key=f"line_{idx}"):
                        send_line_notify(line_token, row['msg_content'])

    except Exception as e:
        st.error(f"發生錯誤：{e}")
else:
    st.info("👆 請上傳 Excel")
