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
    st.header("🕰️ 時光機設定")
    # 預設為電腦的今天，但允許使用者修改 (方便測試 2026 的單)
    simulated_today = st.date_input("設定「今天」日期", datetime.now())
    st.caption(f"目前模擬日期：{simulated_today.strftime('%Y-%m-%d')}")
    
    st.markdown("---")
    st.info("💡 **修正說明：**\n1. 修復變數定義錯誤 (row_res)\n2. 新增日期模擬器 (可測試未來單)\n3. 支援重複標題 (如標的1, 標的1)")

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
st.markdown("### 🚀 時光機回測版 (修復變數錯誤)")

uploaded_file = st.file_uploader("請上傳 Excel (工作表1格式)", type=['xlsx', 'csv']) # 支援 csv 以防萬一

if uploaded_file is not None:
    try:
        # 1. 讀取與清洗
        try:
            # 嘗試讀取 Excel
            df = pd.read_excel(uploaded_file, sheet_name=0, header=0, engine='openpyxl')
        except:
            # 如果失敗，嘗試讀取 CSV (有些使用者的副檔名雖是 xlsx 但內容是 csv)
            try:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file)
            except:
                st.error("❌ 檔案格式讀取失敗，請確認是標準 Excel (.xlsx)")
                st.stop()

        # 處理重複標題造成的 "進場價" 判斷
        # 如果第一列包含數字，可能沒有子標題；如果包含文字"進場價"，則刪除
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
        
        trade_date_idx, _ = find_col_index(cols, ["交易日", "trade date"])
        issue_date_idx, _ = find_col_index(cols, ["發行日", "issue date", "start"])
        final_date_idx, _ = find_col_index(cols, ["最終", "評價", "final", "valuation"])
        maturity_date_idx, _ = find_col_index(cols, ["到期", "maturity", "due date"])
        
        email_idx, _ = find_col_index(cols, ["email", "信箱"])
        name_idx, _ = find_col_index(cols, ["理專", "姓名", "客戶"])

        if t1_idx is None or ko_idx is None:
            st.error("❌ 嚴重錯誤：無法辨識關鍵欄位 (KO 或 標的1)。")
            st.stop()

        # 3. 建立資料表
        clean_df = pd.DataFrame()
        clean_df['ID'] = df.iloc[:, id_idx]
        
        clean_df['TradeDate'] = pd.to_datetime(df.iloc[:, trade_date_idx], errors='coerce') if trade_date_idx else pd.NaT
        clean_df['IssueDate'] = pd.to_datetime(df.iloc[:, issue_date_idx], errors='coerce') if issue_date_idx else pd.Timestamp.min
        clean_df['ValuationDate'] = pd.to_datetime(df.iloc[:, final_date_idx], errors='coerce') if final_date_idx else pd.Timestamp.max
        clean_df['MaturityDate'] = pd.to_datetime(df.iloc[:, maturity_date_idx], errors='coerce') if maturity_date_idx else pd.NaT
        
        def calc_tenure(row):
            if pd.notna(row['MaturityDate']) and pd.notna(row['IssueDate']):
                days = (row['MaturityDate'] - row['IssueDate']).days
                return f"{round(days/30.5, 1)}個月" # 改成月數可能比較直觀，或維持年
            return "-"
        clean_df['Tenure'] = clean_df.apply(calc_tenure, axis=1)

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
                # 如果找不到標的欄位，嘗試用推算的 (假設每2欄一組)
                if tx_idx is None: tx_idx = t1_idx + (i-1)*2
            
            # 防呆：確認 index 沒有超出範圍
            if tx_idx < len(df.columns):
                clean_df[f'T{i}_Code'] = df.iloc[:, tx_idx]
                # 假設下一欄是價格
                if tx_idx + 1 < len(df.columns):
                    clean_df[f'T{i}_Strike'] = df.iloc[:, tx_idx + 1]
                else:
                    clean_df[f'T{i}_Strike'] = 0
            else:
                clean_df[f'T{i}_Code'] = ""
                clean_df[f'T{i}_Strike'] = 0

        clean_df = clean_df.dropna(subset=['ID'])
        
        # 4. 抓取股價
        st.info(f"下載美股資料... (模擬日期: {simulated_today.strftime('%Y-%m-%d')}) ☕")
        all_tickers = []
        for i in range(1, 6):
            if f'T{i}_Code' in clean_df.columns:
                tickers = clean_df[f'T{i}_Code'].dropna().astype(str).unique().tolist()
                all_tickers.extend(tickers)
        all_tickers = [t.strip() for t in set(all_tickers) if t != 'nan' and str(t).strip() != '']
        
        if not all_tickers: st.stop()
            
        min_issue_date = clean_df['IssueDate'].min()
        if pd.isna(min_issue_date): min_issue_date = datetime.now() - timedelta(days=365)
        
        try:
            # 抓取直到模擬日期的資料
            history_data = yf.download(all_tickers, start=min_issue_date, end=simulated_today + timedelta(days=1))['Close']
        except:
            st.error("美股連線失敗")
            st.stop()

        # 5. 核心邏輯
        results = []
        # 使用者設定的模擬今天
        today = pd.Timestamp(simulated_today)

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
                if f'T{i}_Code' not in row: continue
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
            
            # 只取 發行日 ~ 模擬今天
            sim_data = product_history[(product_history.index >= row['IssueDate']) & (product_history.index <= today)]
            
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
            
            # --- 狀態計算 ---
            locked_list = []
            waiting_list = []
            hit_ki_list = []
            detail_cols = {}

            # 取得「模擬今天」的最新價格
            for i, asset in enumerate(assets):
                try:
                    # 嘗試抓最後一筆 (如果 sim_data 空的，表示還沒發行或沒資料)
                    if not sim_data.empty:
                        if len(all_tickers) == 1: curr = float(sim_data.iloc[-1])
                        else: curr = float(sim_data.iloc[-1][asset['code']])
                    else:
                        curr = 0 # 沒資料
                    
                    if curr > 0:
                        asset['price'] = curr
                        asset['perf'] = curr / asset['initial']
                        if not is_aki and asset['perf'] < ki_thresh: 
                            asset['hit_ki'] = True
                            asset['ki_record'] = f"@{curr:.2f} (EKI)"
                except: pass
                
                if asset['locked_ko']: locked_list.append(asset['code'])
                else: waiting_list.append(asset['code'])
                if asset['hit_ki']: hit_ki_list.append(asset['code'])
                
                p_pct = round(asset['perf']*100, 2)
                status_icon = "✅" if asset['locked_ko'] else "⚠️" if asset['hit_ki'] else ""
                
                cell_text = f"{asset['code']}\n${round(asset['price'], 2)} / ${round(asset['initial'], 2)}\n{p_pct}% {status_icon}"
                if asset['locked_ko']: cell_text += f"\nKO {asset['ko_record']}"
                if asset['hit_ki']: cell_text += f"\nKI {asset['ki_record']}"
                
                detail_cols[f"T{i+1}_Detail"] = cell_text

            hit_any_ki = any(a['hit_ki'] for a in assets)
            all_above_strike_now = all(a['perf'] >= strike_thresh for a in assets)
            
            # 防呆：如果 assets 裡有人 perf 是 0 (沒抓到股價)，min 會錯
            valid_assets = [a for a in assets if a['perf'] > 0]
            if valid_assets:
                worst_asset = min(valid_assets, key=lambda x: x['perf'])
                worst_perf = worst_asset['perf']
                worst_code = worst_asset['code']
                worst_strike_price = worst_asset['strike_price']
            else:
                worst_perf = 0
                worst_code = "N/A"
                worst_strike_price = 0
            
            # --- 總結狀態 ---
            final_status = ""
            if today < row['IssueDate']:
                final_status = "⏳ 未發行"
            elif product_status == "Early Redemption":
                final_status = f"🎉 提前出場\n({early_redemption_date.strftime('%Y-%m-%d')})"
            elif pd.notna(row['ValuationDate']) and today >= row['ValuationDate']:
                if all_above_strike_now:
                     final_status = "💰 到期獲利\n(全數 > 執行價)"
                elif hit_any_ki:
                     final_status = f"😭 到期接股\n{worst_code} @ {round(worst_strike_price, 2)}"
                else:
                     final_status = "🛡️ 到期保本\n(未破KI)"
            else:
                if not waiting_list:
                    final_status = "👀 比價中"
                else:
                    wait_str = ",".join(waiting_list)
                    final_status = f"👀 比價中\n⏳等待: {wait_str}"
                    if locked_list:
                         final_status += f"\n✅已鎖: {','.join(locked_list)}"
                
                if hit_any_ki:
                    final_status += f"\n⚠️ KI已破: {','.join(hit_ki_list)}"

            # 格式化日期字串
            trade_date_str = row['TradeDate'].strftime('%Y-%m-%d') if pd.notna(row['TradeDate']) else "-"
            issue_date_str = row['IssueDate'].strftime('%Y-%m-%d') if pd.notna(row['IssueDate']) else "-"
            val_date_str = row['ValuationDate'].strftime('%Y-%m-%d') if pd.notna(row['ValuationDate']) else "-"
            mat_date_str = row['MaturityDate'].strftime('%Y-%m-%d') if pd.notna(row['MaturityDate']) else "-"

            # --- 組合結果 ---
            # 這裡先定義 row_res，確保後面不會有變數未定義的問題
            row_res = {
                "債券代號": row['ID'],
                "天期": row['Tenure'],
                "收件人": row['Name'],
                "Email": str(row['Email']).strip(),
                "狀態": final_status,
                "最差表現": f"{round(worst_perf*100, 2)}%",
                "KO設定": f"{ko_thresh_val}%",
                "KI設定": f"{ki_thresh_val}%",
                "執行價": f"{strike_thresh_val}%",
                
                "交易日": trade_date_str,
                "發行日": issue_date_str,
                "最終評價": val_date_str,
                "到期日": mat_date_str,
                
                "msg_subject": f"【ELN通知】{row['ID']} 狀態更新",
                "msg_body": (
                    f"Hi {row['Name']}：\n\n"
                    f"商品 {row['ID']} 最新報告：\n"
                    f"📊 狀態：\n{final_status}\n\n"
                    f"⚡ 設定：KO {ko_thresh_val}% / KI {ki_thresh_val}% ({row['KI_Type']})\n"
                    f"📉 執行價格(Strike)：{strike_thresh_val}%\n"
                    f"📅 到期日：{mat_date_str}\n\n"
                    f"--------------------------------\n"
                    f"系統自動發送"
                )
            }
            # 最後再合併詳細資訊
            row_res.update(detail_cols)
            results.append(row_res)

        # 6. 顯示
        if not results:
            st.warning("⚠️ 沒有讀取到有效的資料列，請檢查 Excel 內容。")
        else:
            final_df = pd.DataFrame(results)
            
            st.subheader("📋 專業監控列表")
            
            def color_status(val):
                if "提前" in str(val) or "獲利" in str(val): return 'background-color: #d4edda; color: green'
                if "接股" in str(val) or "KI" in str(val): return 'background-color: #f8d7da; color: red'
                if "未發行" in str(val): return 'background-color: #fff3cd; color: #856404'
                return ''

            # 動態取得標的欄位
            t_cols = [c for c in final_df.columns if '_Detail' in c]
            t_cols.sort()
            
            display_cols = ['債券代號', '天期', '狀態', '最差表現', 'KO設定', 'KI設定', '執行價'] + \
                           t_cols + \
                           ['交易日', '發行日', '最終評價', '到期日']
            
            # 建立表格設定
            column_config = {
                "狀態": st.column_config.TextColumn("目前狀態摘要", width="large"),
                "債券代號": st.column_config.TextColumn("代號", width="small"),
                "天期": st.column_config.TextColumn("天期", width="small"),
                "KO設定": st.column_config.TextColumn("KO", width="small"),
                "KI設定": st.column_config.TextColumn("KI", width="small"),
                "執行價": st.column_config.TextColumn("Strike", width="small"),
                "最差表現": st.column_config.TextColumn("Worst Of", width="small"),
            }
            for i, c in enumerate(t_cols):
                column_config[c] = st.column_config.TextColumn(f"標的 {i+1}", width="medium")

            st.dataframe(
                final_df[display_cols].style.applymap(color_status, subset=['狀態']), 
                use_container_width=True, 
                column_config=column_config,
                height=600,
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
