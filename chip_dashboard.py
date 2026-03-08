import streamlit as st
import pandas as pd
import requests
import urllib3
import time
import os

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def fetch_full_market_data(date_str, target_stocks):
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

    # --- 引擎 A：抓籌碼 (法人動向) ---
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
    except:
        return None

    # --- 引擎 B：狙擊手抓價量 ---
    roc_year = int(date_str[:4]) - 1911
    roc_date_str = f"{roc_year}/{date_str[4:6]}/{date_str[6:8]}"
    price_list = []
    
    for stock in target_stocks:
        url_price = f"https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date={date_str}&stockNo={stock}"
        try:
            res = requests.get(url_price, headers=headers, verify=False, timeout=5)
            data = res.json()
            if data.get('stat') == 'OK':
                for row in data['data']:
                    if row[0] == roc_date_str:
                        vol_int = int(int(row[1].replace(',', '')) / 1000)
                        close_price = float(row[6].replace(',', ''))
                        price_list.append({'代號': stock, '收盤價': close_price, '總成交量(張)': vol_int})
                        break
        except:
            pass
        time.sleep(1) 
    df_price = pd.DataFrame(price_list)

    # --- 引擎 C：抓散戶指標 (融資餘額) ---
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
    except:
        pass # 若抓不到散戶資料，不影響主程式運行

    # --- 完美縫合 ---
    if not df_price.empty:
        df_final = pd.merge(df_chips, df_price, on='代號', how='left')
        if not df_margin.empty:
            df_final = pd.merge(df_final, df_margin, on='代號', how='left')
        else:
            df_final['融資餘額(張)'] = 0
            
        df_final['法人買超佔比(%)'] = df_final.apply(
            lambda row: round((row['三大法人合計(張)'] / row['總成交量(張)']) * 100, 2) if row['總成交量(張)'] > 0 else 0, axis=1)
        return df_final
    else:
        st.warning(f"⚠️ 找不到 {date_str} 的價量資料。")
        return df_chips

# ================= 網頁介面與邏輯 =================
st.set_page_config(page_title="專屬籌碼戰情室", layout="wide")
st.title("🎯 專屬籌碼分析戰情室 (含散戶反向雷達)")

if 'current_data' not in st.session_state:
    st.session_state.current_data, st.session_state.current_date = None, None

st.sidebar.header("⚙️ 戰略設定")
current_dir = os.path.dirname(os.path.abspath(__file__))
list_file = os.path.join(current_dir, "my_stocks.txt")
history_file = os.path.join(current_dir, "chip_history.csv")

if os.path.exists(list_file):
    with open(list_file, "r", encoding="utf-8") as f:
        default_stocks = f.read().strip()
else:
    default_stocks = "1513, 1514, 2886, 1216, 9904"

stock_input = st.sidebar.text_input("觀察清單 (代號逗號分隔)", default_stocks)
my_stocks = [s.strip() for s in stock_input.split(',')]
selected_date = st.sidebar.date_input("選擇交易日").strftime("%Y%m%d")

if st.sidebar.button("🔍 執行籌碼掃描"):
    with open(list_file, "w", encoding="utf-8") as f: f.write(stock_input)
    with st.spinner('🎯 狙擊手出動，抓取三大法人與散戶資料中...'):
        df = fetch_full_market_data(selected_date, my_stocks)
        if df is not None:
            st.session_state.current_data = df
            st.session_state.current_date = selected_date

if st.session_state.current_data is not None:
    df_show = st.session_state.current_data.copy()
    current_d = st.session_state.current_date
    st.success(f"✅ 成功獲取 {current_d} 數據！")
    
    # 計算連續天數
    if os.path.exists(history_file):
        hist_df = pd.read_csv(history_file)
        hist_df = hist_df[hist_df['日期'].astype(str) != str(current_d)] 
        temp_curr = df_show.copy()
        temp_curr.insert(0, '日期', current_d)
        full_df = pd.concat([hist_df, temp_curr], ignore_index=True)
        full_df['日期'] = full_df['日期'].astype(str)
        full_df = full_df.sort_values('日期', ascending=False)
        
        s_list, t_list = [], []
        for stock in df_show['名稱']:
            for data_list, res_list in [(full_df[full_df['名稱'] == stock]['三大法人合計(張)'].tolist(), s_list), 
                                        (full_df[full_df['名稱'] == stock]['投信買超(張)'].tolist(), t_list)]:
                streak = 0
                if len(data_list) > 0:
                    first = data_list[0]
                    if first > 0:
                        for v in data_list:
                            if v > 0: streak += 1
                            else: break
                        res_list.append(f"🔴 連買 {streak} 天")
                    elif first < 0:
                        for v in data_list:
                            if v < 0: streak += 1
                            else: break
                        res_list.append(f"🟢 連賣 {streak} 天")
                    else: res_list.append("⚪ 平盤")
                else: res_list.append("⚪ 無資料")
        df_show['法人動向'], df_show['投信動向'] = s_list, t_list
    else:
        df_show['法人動向'] = df_show['投信動向'] = "📝 需存檔"

    cols = ['代號', '名稱', '收盤價', '投信動向', '法人動向', '法人買超佔比(%)', '融資餘額(張)', '總成交量(張)', '外資買超(張)', '投信買超(張)', '三大法人合計(張)']
    df_show = df_show[[c for c in cols if c in df_show.columns]]

    col1, col2 = st.columns([1.5, 1])
    with col1:
        st.markdown("### 📋 綜合數據總表")
        st.dataframe(df_show, use_container_width=True)
    with col2:
        st.markdown("### 📊 法人買超比較")
        chart_data = df_show.set_index('名稱')[['外資買超(張)', '投信買超(張)']]
        st.bar_chart(chart_data)

    st.markdown("### 💾 資料庫管理")
    # === 💡 新增：一鍵下載按鈕 ===
    # 將目前的表格轉換成 CSV 格式，並加上 utf-8-sig 確保 Excel 打開中文不會亂碼
    csv_data = df_show.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="⬇️ 一鍵下載今日數據 (CSV/Excel)",
        data=csv_data,
        file_name=f"籌碼戰況_{current_d}.csv",
        mime="text/csv"
    )
    # ==============================

    if st.button("📥 將本日數據存入歷史資料庫"):
        df_to_save = df_show.copy()
        # 清除圖表用的動向文字，保持資料庫純淨
        if '法人動向' in df_to_save.columns: df_to_save = df_to_save.drop(columns=['法人動向', '投信動向'])
        df_to_save.insert(0, '日期', current_d) 
        if not os.path.exists(history_file):
            df_to_save.to_csv(history_file, index=False, encoding='utf-8-sig')
            st.success("🎉 歷史資料庫建立成功！")
        else:
            history_df = pd.read_csv(history_file)
            if str(current_d) in history_df['日期'].astype(str).values:
                history_df = history_df[history_df['日期'].astype(str) != str(current_d)]
            updated_df = pd.concat([history_df, df_to_save], ignore_index=True)
            updated_df.to_csv(history_file, index=False, encoding='utf-8-sig')
            st.success(f"✅ {current_d} 數據已成功更新至資料庫！")

    st.markdown("---")
    st.markdown("### 📈 歷史籌碼與股價趨勢分析")
    if os.path.exists(history_file):
        df_hist = pd.read_csv(history_file)
        df_hist['日期'] = df_hist['日期'].astype(str)
        df_hist = df_hist.sort_values('日期')
        avail_stocks = df_hist['名稱'].unique().tolist()
        if avail_stocks:
            sel_stock = st.selectbox("請選擇要分析的股票：", avail_stocks)
            df_st_hist = df_hist[df_hist['名稱'] == sel_stock].set_index('日期')
            c3, c4 = st.columns(2)
            with c3:
                st.markdown(f"**{sel_stock} - 收盤價**")
                st.line_chart(df_st_hist['收盤價'])
            with c4:
                st.markdown(f"**{sel_stock} - 三大法人與散戶(融資)**")
                chart_cols = ['三大法人合計(張)']
                if '融資餘額(張)' in df_st_hist.columns: chart_cols.append('融資餘額(張)')

                st.bar_chart(df_st_hist[chart_cols])



