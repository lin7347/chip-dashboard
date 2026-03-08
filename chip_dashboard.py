import streamlit as st
import pandas as pd
import requests
import urllib3
import time
import json
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ================= 雲端大腦連線設定 =================
@st.cache_resource
def init_connection():
    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        raw_secret = st.secrets["google_credentials"]
        
        if isinstance(raw_secret, str):
            try:
                creds_dict = json.loads(raw_secret, strict=False)
            except:
                clean_secret = raw_secret.replace('\n', '\\n').replace('\r', '')
                creds_dict = json.loads(clean_secret, strict=False)
        else:
            creds_dict = dict(raw_secret)

        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client.open("專屬籌碼資料庫")
    except Exception as e:
        st.error(f"❌ 無法連線至 Google 試算表：{e}")
        return None

sheet = init_connection()

# ================= 資料抓取引擎 =================
def fetch_full_market_data(date_str, target_stocks):
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

    # --- 引擎 A：抓籌碼 ---
    url_chips = f"https://www.twse.com.tw/fund/T86?response=json&date={date_str}&selectType=ALL"
    try:
        res_chips = requests.get(url_chips, headers=headers, verify=False, timeout=10)
        data_chips = res_chips.json()
        if data_chips.get('stat') != 'OK' or 'data' not in data_chips: return None
            
        df_chips = pd.DataFrame(data_chips['data'], columns=data_chips['fields'])
        df_chips = df_chips[df_chips['證券代號'].isin(target_stocks)][['證券代號', '證券名稱', '外陸資買賣超股數(不含外資自營商)', '投信買賣超股數', '三大法人買賣超股數']]
        df_chips.columns = ['代號', '名稱', '外資買超(張)', '投信買超(張)', '三大法人合計(張)']
        for col in ['外資買超(張)', '投信買超(張)', '三大法人合計(張)']:
            df_chips[col] = df_chips[col].astype(str).str.replace(',', '', regex=False).astype(float) / 1000
            df_chips[col] = df_chips[col].astype(int)
    except: return None

    # --- 引擎 B：抓價量與「5日均量判定」 ---
    roc_year = int(date_str[:4]) - 1911
    roc_date_str = f"{roc_year}/{date_str[4:6]}/{date_str[6:8]}"
    price_list = []
    for stock in target_stocks:
        url_price = f"https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date={date_str}&stockNo={stock}"
        try:
            res = requests.get(url_price, headers=headers, verify=False, timeout=5)
            data = res.json()
            if data.get('stat') == 'OK':
                records = data['data']
                target_idx = -1
                for i, row in enumerate(records):
                    if row[0] == roc_date_str:
                        target_idx = i
                        break
                
                if target_idx != -1:
                    vol_int = int(int(records[target_idx][1].replace(',', '')) / 1000)
                    close_price = float(records[target_idx][6].replace(',', ''))
                    
                    # 計算均量 (最多回溯當月前4個交易日，加上當日共5日)
                    start_idx = max(0, target_idx - 4)
                    vol_sum = 0
                    count = 0
                    for j in range(start_idx, target_idx + 1):
                        vol_sum += int(int(records[j][1].replace(',', '')) / 1000)
                        count += 1
                    avg_vol = round(vol_sum / count) if count > 0 else 0
                    
                    # 判斷量能狀態
                    if vol_int > avg_vol * 1.2: vol_status = "🔥 爆量"
                    elif vol_int < avg_vol * 0.8: vol_status = "💧 量縮"
                    else: vol_status = "⚪ 量平"

                    price_list.append({'代號': stock, '收盤價': close_price, '總成交量(張)': vol_int, '5日均量(張)': avg_vol, '量能變化': vol_status})
        except: pass
        time.sleep(1) 
    df_price = pd.DataFrame(price_list)

    # --- 引擎 C：抓散戶指標 (融資) ---
    url_margin = f"https://www.twse.com.tw/exchangeReport/MI_MARGN?response=json&date={date_str}&selectType=ALL"
    df_margin = pd.DataFrame()
    try:
        res_margin = requests.get(url_margin, headers=headers, verify=False, timeout=10)
        data_margin = res_margin.json()
        if data_margin.get('stat') == 'OK':
            target_data, target_fields = None, None
            if 'tables' in data_margin:
                for table in data_margin['tables']:
                    if 'fields' in table and any('融資餘額' in str(f) for f in table['fields']):
                        target_data, target_fields = table['data'], table['fields']
                        break
            else:
                for key, val in data_margin.items():
                    if key.startswith('fields') and any('融資餘額' in str(f) for f in val):
                        data_key = key.replace('fields', 'data')
                        target_data, target_fields = data_margin.get(data_key), val
                        break
            if target_data and target_fields:
                temp_margin = pd.DataFrame(target_data, columns=target_fields)
                code_col = [c for c in temp_margin.columns if '代號' in c][0]
                bal_col = [c for c in temp_margin.columns if '融資餘額' in c][0]
                df_margin = temp_margin[temp_margin[code_col].isin(target_stocks)][[code_col, bal_col]]
                df_margin.columns = ['代號', '融資餘額(張)']
                df_margin['融資餘額(張)'] = df_margin['融資餘額(張)'].astype(str).str.replace(',', '', regex=False).astype(int)
    except: pass 

    # --- 完美縫合 ---
    if not df_price.empty:
        df_final = pd.merge(df_chips, df_price, on='代號', how='left')
        if not df_margin.empty: df_final = pd.merge(df_final, df_margin, on='代號', how='left')
        else: df_final['融資餘額(張)'] = 0
        df_final['法人買超佔比(%)'] = df_final.apply(lambda row: round((row['三大法人合計(張)'] / row['總成交量(張)']) * 100, 2) if row['總成交量(張)'] > 0 else 0, axis=1)
        return df_final
    else:
        st.warning(f"⚠️ 找不到 {date_str} 的價量資料。")
        return df_chips

# ================= 網頁介面與邏輯 =================
st.set_page_config(page_title="專屬籌碼戰情室", layout="wide")
st.title("🎯 專屬籌碼分析戰情室 (量能擴充版)")

if 'current_data' not in st.session_state:
    st.session_state.current_data, st.session_state.current_date = None, None

st.sidebar.header("⚙️ 戰略設定")

# --- 讀取 Google 觀察清單 ---
default_stocks_str = "1513, 1514, 2886, 1216, 9904"
if sheet:
    try:
        ws_list = sheet.worksheet("觀察清單")
        records = ws_list.col_values(1)
        if len(records) > 1:
            default_stocks_str = ", ".join([str(x) for x in records[1:] if str(x).strip() != ''])
    except: pass

stock_input = st.sidebar.text_input("觀察清單 (代號逗號分隔)", default_stocks_str)
my_stocks = [s.strip() for s in stock_input.split(',') if s.strip()]
selected_date = st.sidebar.date_input("選擇交易日").strftime("%Y%m%d")

if st.sidebar.button("🔍 執行籌碼掃描"):
    if sheet:
        try:
            ws_list = sheet.worksheet("觀察清單")
            ws_list.clear()
            ws_list.update('A1', [['代號']] + [[s] for s in my_stocks])
        except: pass

    with st.spinner('🎯 雲端大腦啟動，跨海抓取資料中...'):
        df = fetch_full_market_data(selected_date, my_stocks)
        if df is not None:
            st.session_state.current_data = df
            st.session_state.current_date = selected_date

# --- 讀取 Google 歷史資料庫 ---
df_hist = pd.DataFrame()
if sheet:
    try:
        ws_hist = sheet.worksheet("歷史數據")
        hist_records = ws_hist.get_all_records()
        if hist_records:
            df_hist = pd.DataFrame(hist_records)
            df_hist['日期'] = df_hist['日期'].astype(str)
    except: pass

# --- 顯示資料區 ---
if st.session_state.current_data is not None:
    df_show = st.session_state.current_data.copy()
    current_d = st.session_state.current_date
    st.success(f"✅ 成功獲取 {current_d} 數據！")
    
    if not df_hist.empty:
        hist_df_filtered = df_hist[df_hist['日期'] != str(current_d)] 
        temp_curr = df_show.copy()
        temp_curr.insert(0, '日期', current_d)
        full_df = pd.concat([hist_df_filtered, temp_curr], ignore_index=True)
        full_df['日期'] = full_df['日期'].astype(str)
        full_df = full_df.sort_values('日期', ascending=False)
        
        s_list, t_list = [], []
        for stock in df_show['名稱']:
            for data_list, res_list in [(full_df[full_df['名稱'] == stock]['三大法人合計(張)'].tolist(), s_list), 
                                        (full_df[full_df['名稱'] == stock]['投信買超(張)'].tolist(), t_list)]:
                streak = 0
                if len(data_list) > 0:
                    first = float(data_list[0]) if pd.notna(data_list[0]) and str(data_list[0]).strip()!='' else 0
                    if first > 0:
                        for v in data_list:
                            v_float = float(v) if pd.notna(v) and str(v).strip()!='' else 0
                            if v_float > 0: streak += 1
                            else: break
                        res_list.append(f"🔴 連買 {streak} 天")
                    elif first < 0:
                        for v in data_list:
                            v_float = float(v) if pd.notna(v) and str(v).strip()!='' else 0
                            if v_float < 0: streak += 1
                            else: break
                        res_list.append(f"🟢 連賣 {streak} 天")
                    else: res_list.append("⚪ 平盤")
                else: res_list.append("⚪ 無資料")
        df_show['法人動向'], df_show['投信動向'] = s_list, t_list
    else:
        df_show['法人動向'] = df_show['投信動向'] = "📝 需存檔"

    # 加入量能變化的欄位排序
    cols = ['代號', '名稱', '收盤價', '投信動向', '法人動向', '量能變化', '總成交量(張)', '5日均量(張)', '法人買超佔比(%)', '融資餘額(張)', '外資買超(張)', '投信買超(張)', '三大法人合計(張)']
    df_show = df_show[[c for c in cols if c in df_show.columns]]

    # --- 第一區：核心看板 ---
    st.markdown("### 📋 綜合數據總表")
    st.dataframe(df_show, hide_index=True, use_container_width=True)
    st.markdown("---")
    st.markdown("### 📊 法人買超比較")
    chart_data = df_show.set_index('名稱')[['外資買超(張)', '投信買超(張)']]
    st.bar_chart(chart_data, height=400)
    st.markdown("---")

    # --- 第二區：資料庫管理抽屜 ---
    with st.expander("💾 資料庫管理與下載 (點擊展開)"):
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            csv_data = df_show.to_csv(index=False).encode('utf-8-sig')
            st.download_button(label="⬇️ 下載今日數據 (Excel)", data=csv_data, file_name=f"籌碼戰況_{current_d}.csv", mime="text/csv")
            
        with col_btn2:
            if st.button("📥 寫入 Google 雲端歷史資料庫"):
                if sheet:
                    df_to_save = df_show.copy()
                    if '法人動向' in df_to_save.columns: df_to_save = df_to_save.drop(columns=['法人動向', '投信動向', '量能變化'])
                    df_to_save.insert(0, '日期', current_d)
                    
                    try:
                        ws_hist = sheet.worksheet("歷史數據")
                        if not df_hist.empty:
                            hist_df_filtered = df_hist[df_hist['日期'] != str(current_d)]
                            updated_df = pd.concat([hist_df_filtered, df_to_save], ignore_index=True)
                        else:
                            updated_df = df_to_save
                        
                        updated_df = updated_df.astype(str)
                        ws_hist.clear()
                        ws_hist.update('A1', [updated_df.columns.values.tolist()] + updated_df.values.tolist())
                        st.success(f"✅ {current_d} 數據已成功寫入您的 Google 試算表！")
                    except Exception as e:
                        st.error(f"❌ 寫入失敗：{e}")

    # --- 第三區：歷史趨勢分析 ---
    with st.expander("📈 歷史籌碼與股價趨勢分析 (點擊展開)"):
        if not df_hist.empty:
            df_hist_display = df_hist.copy()
            df_hist_display['日期'] = df_hist_display['日期'].astype(str)
            df_hist_display = df_hist_display.sort_values('日期')
            avail_stocks = df_hist_display['名稱'].unique().tolist()
            if avail_stocks:
                sel_stock = st.selectbox("請選擇要分析的股票：", avail_stocks)
                df_st_hist = df_hist_display[df_hist_display['名稱'] == sel_stock].set_index('日期')
                
                for col in ['收盤價', '三大法人合計(張)', '融資餘額(張)']:
                    if col in df_st_hist.columns:
                        df_st_hist[col] = pd.to_numeric(df_st_hist[col], errors='coerce').fillna(0)
                        
                c3, c4 = st.columns(2)
                with c3:
                    st.markdown(f"**{sel_stock} - 收盤價**")
                    st.line_chart(df_st_hist['收盤價'])
                with c4:
                    st.markdown(f"**{sel_stock} - 三大法人與散戶(融資)**")
                    chart_cols = ['三大法人合計(張)']
                    if '融資餘額(張)' in df_st_hist.columns: chart_cols.append('融資餘額(張)')
                    st.bar_chart(df_st_hist[chart_cols])
        else:
            st.info("📝 尚未有歷史紀錄，請先執行上方掃描並存檔。")

    # --- 第四區：區間籌碼累計加總 (全新功能！) ---
    with st.expander("📅 區間籌碼累計加總 (點擊展開)"):
        if not df_hist.empty:
            st.markdown("計算所選日期區間內，觀察清單股票的三大法人累計買賣超張數。")
            df_hist_dates = pd.to_datetime(df_hist['日期'], format='%Y%m%d', errors='coerce').dropna()
            
            if not df_hist_dates.empty:
                min_date, max_date = df_hist_dates.min().date(), df_hist_dates.max().date()
                date_range = st.date_input("選擇歷史分析區間", [min_date, max_date], min_value=min_date, max_value=max_date)
                
                if len(date_range) == 2:
                    start_str, end_str = date_range[0].strftime("%Y%m%d"), date_range[1].strftime("%Y%m%d")
                    mask = (df_hist['日期'] >= start_str) & (df_hist['日期'] <= end_str)
                    df_period = df_hist.loc[mask].copy()

                    if not df_period.empty:
                        for col in ['外資買超(張)', '投信買超(張)', '三大法人合計(張)']:
                            if col in df_period.columns:
                                df_period[col] = pd.to_numeric(df_period[col], errors='coerce').fillna(0)
                        
                        df_agg = df_period.groupby(['代號', '名稱'])[['外資買超(張)', '投信買超(張)', '三大法人合計(張)']].sum().reset_index()
                        st.dataframe(df_agg, hide_index=True, use_container_width=True)
                    else:
                        st.info("該區間內沒有歷史數據，請確認您的 Google 試算表中是否有此區間的存檔。")
        else:
            st.info("📝 您的 Google 試算表目前是空的，只要每天按一次存檔，這裡就會自動幫您計算區間總和喔！")

# --- 第五區：完整歷史資料庫瀏覽 (查看試算表全部內容) ---
    with st.expander("🗄️ 雲端完整歷史資料庫 (點擊展開)"):
        if not df_hist.empty:
            st.markdown("這裡顯示您存在 Google 試算表中的所有 30 檔 (或更多) 股票的歷史紀錄。")
            
            # 製作一個下拉選單，讓您可以看全部，或是快速篩選某一檔
            all_stocks_in_db = ["顯示全部"] + list(df_hist['名稱'].unique())
            filter_stock = st.selectbox("篩選特定股票歷史紀錄：", all_stocks_in_db, key="db_filter")
            
            if filter_stock == "顯示全部":
                st.dataframe(df_hist, hide_index=True, use_container_width=True)
            else:
                st.dataframe(df_hist[df_hist['名稱'] == filter_stock], hide_index=True, use_container_width=True)
        else:
            st.info("📝 您的 Google 試算表目前是空的。")
