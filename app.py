import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import urllib.request
from matplotlib import font_manager

# ----------------- 0. 強制載入中文字型 (終極解決方塊字) -----------------
@st.cache_resource
def load_font():
    import requests
    font_url = "https://raw.githubusercontent.com/StellarCN/scp_zh/master/fonts/SimHei.ttf"
    font_path = "SimHei.ttf"
    
    if not os.path.exists(font_path):
        response = requests.get(font_url)
        with open(font_path, 'wb') as f:
            f.write(response.content)
            
    font_manager.fontManager.addfont(font_path)
    prop = font_manager.FontProperties(fname=font_path)
    
    plt.rcParams['font.sans-serif'] = [prop.get_name(), 'sans-serif']
    plt.rcParams['axes.unicode_minus'] = False

load_font()

# ----------------- 1. 網頁基本設定 -----------------
st.set_page_config(page_title="航班起降統計報表產生器", page_icon="✈️", layout="wide")

# 隱藏右上角 GitHub 連結與頂部工具列
st.markdown("""
    <style>
    [data-testid="stHeader"] {
        display: none;
    }
    </style>
""", unsafe_allow_html=True)

# 【左右分欄設計微調】調整比例讓標題與 Logo 距離拉近，Logo 靠右對齊
col_title, col_logo = st.columns([6, 1])

with col_title:
    st.title("✈️ 航班統計報表產生器 (網頁版)")
    st.markdown("**2026.10 V1.0**")

with col_logo:
    if os.path.exists("logo.png"):
        # 稍微縮小寬度並靠右
        st.markdown("<div style='text-align: right;'>", unsafe_allow_html=True)
        st.image("logo.png", width=180)
        st.markdown("</div>", unsafe_allow_html=True)
def get_hour(time_val):
    if pd.isna(time_val): return None
    time_str = str(time_val).strip()
    if ' ' in time_str: time_str = time_str.split(' ')[-1]
    if len(time_str) >= 2: return time_str[:2]
    return None

# ----------------- 2. 檔案上傳與處理區 -----------------
uploaded_file = st.file_uploader("📂 請拖曳或點擊選擇 Excel 檔案 (.xls, .xlsx)", type=['xls', 'xlsx'])

if uploaded_file is not None:
    with st.spinner("⚙️ 正在分析航班資料..."):
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = df.columns.str.strip()
            
           # 判斷基準日期
            if '班機日期' in df.columns:
                target_date_str = str(df['班機日期'].mode()[0])[:10]
            elif 'ETA(D)' in df.columns:
                target_date_str = str(df['ETA(D)'].dropna().mode()[0])[:10]
            else:
                target_date_str = '2026-10-24'
                
            # 【核心修正】強力清理班號欄位：徹底清除隱形空白、\xa0、nan，確保與單機版過濾邏輯完全一致
            def clean_flight_no(val):
                if pd.isna(val): return ''
                s = str(val).replace('\xa0', '').strip()
                if s.lower() in ['nan', 'none', '']: return ''
                return s

            df['入境_clean'] = df['入境'].apply(clean_flight_no) if '入境' in df.columns else ''
            df['出境_clean'] = df['出境'].apply(clean_flight_no) if '出境' in df.columns else ''
                
            df['ETA_Date_Str'] = df['ETA(D)'].astype(str).str[:10] if 'ETA(D)' in df.columns else target_date_str
            df['ETD_Date_Str'] = df['ETD(D)'].astype(str).str[:10] if 'ETD(D)' in df.columns else target_date_str
            df['Arr_Hour'] = df['ETA'].apply(get_hour) if 'ETA' in df.columns else None
            df['Dep_Hour'] = df['ETD'].apply(get_hour) if 'ETD' in df.columns else None
            
            # 預計起降防呆過濾
            mask_eta = (df['ETA_Date_Str'] == target_date_str) & (df['入境_clean'] != '')
            mask_etd = (df['ETD_Date_Str'] == target_date_str) & (df['出境_clean'] != '')
            
            valid_arr = df.loc[mask_eta, 'Arr_Hour'] if df['Arr_Hour'] is not None else pd.Series()
            valid_dep = df.loc[mask_etd, 'Dep_Hour'] if df['Dep_Hour'] is not None else pd.Series()
            
            arr_counts = valid_arr.value_counts().to_dict()
            dep_counts = valid_dep.value_counts().to_dict()
            
            hours = [f"{i:02d}" for i in range(24)]
            labels = [f"{i:02d}:00" for i in range(24)]
            
            eta_vals = [arr_counts.get(h, 0) for h in hours]
            etd_vals = [dep_counts.get(h, 0) for h in hours]
            sch_total_vals = [eta_vals[i] + etd_vals[i] for i in range(24)]
            
            has_actual = False
            act_total_vals = [np.nan] * 24
            act_table_total = [0] * 24
            ata_vals = [0] * 24
            atd_vals = [0] * 24
            
            if 'ATA' in df.columns and 'ATD' in df.columns:
                df['ATA_Date_Str'] = df['ATA(D)'].astype(str).str[:10] if 'ATA(D)' in df.columns else target_date_str
                df['ATD_Date_Str'] = df['ATD(D)'].astype(str).str[:10] if 'ATD(D)' in df.columns else target_date_str
                df['ATA_Hour'] = df['ATA'].apply(get_hour)
                df['ATD_Hour'] = df['ATD'].apply(get_hour)
                
                # 實際起降防呆過濾
                mask_ata = (df['ATA_Date_Str'] == target_date_str) & (df['入境_clean'] != '')
                mask_atd = (df['ATD_Date_Str'] == target_date_str) & (df['出境_clean'] != '')
                
                valid_ata = df.loc[mask_ata, 'ATA_Hour']
                valid_atd = df.loc[mask_atd, 'ATD_Hour']
                
                ata_counts = valid_ata.value_counts().to_dict()
                atd_counts = valid_atd.value_counts().to_dict()
                
                ata_vals = [ata_counts.get(h, 0) for h in hours]
                atd_vals = [atd_counts.get(h, 0) for h in hours]
                act_table_total = [ata_vals[i] + atd_vals[i] for i in range(24)]
                
                last_actual_hour = -1
                for i in range(24):
                    if ata_vals[i] > 0 or atd_vals[i] > 0:
                        last_actual_hour = i
                        
                if last_actual_hour >= 0:
                    has_actual = True
                    for i in range(24):
                        if i <= last_actual_hour:
                            act_total_vals[i] = act_table_total[i]
                            
            st.success(f"✅ 資料分析完成！基準日期為：{target_date_str}")
            
            # ----------------- 3. 繪圖區 -----------------
            def create_table_fig(vals_arr, vals_dep, vals_tot, title):
                t_data = []
                for i, h in enumerate(hours):
                    next_h = f"{(int(h) + 1):02d}"
                    if next_h == "24": next_h = "00"
                    t_data.append([f"{h}-{next_h}", vals_arr[i], vals_dep[i], vals_tot[i]])
                
                t_arr, t_dep, t_all = sum(vals_arr), sum(vals_dep), sum(vals_tot)
                t_data.append(["總計", f"到場 {t_arr}", f"離場 {t_dep}", f"總計 {t_all}"])
                
                cols = ["時段", "到場", "離場", "總計"]
                f_tbl, a_tbl = plt.subplots(figsize=(6, 11))
                a_tbl.axis('off')
                
                a_tbl.set_title(title, fontsize=16, weight='bold', color='#142952', pad=15)
                tbl = a_tbl.table(cellText=t_data, colLabels=cols, loc='center', cellLoc='center', bbox=[0, 0, 1, 1])
                tbl.auto_set_font_size(False)
                tbl.set_fontsize(14)
                
                for j in range(len(cols)):
                    c = tbl[0, j]
                    c.set_text_props(weight='bold', color='black', fontsize=15)
                    c.set_facecolor('#e0e0e0')

                for i in range(1, len(t_data) + 1):
                    for j in range(len(cols)):
                        c = tbl[i, j]
                        if i == len(t_data):
                            c.set_text_props(weight='bold', fontsize=15)
                            c.set_facecolor('#fff2cc')
                        else:
                            c.set_facecolor('#f4f8fb' if i % 2 == 0 else '#ffffff')
                            if j == 1: c.set_text_props(color='#1e8c45', weight='bold')
                            elif j == 2: c.set_text_props(color='#1c3c78', weight='bold')
                            elif j == 3: c.set_text_props(color='#961536', weight='bold')
                        if j == 0 and i != len(t_data):
                            c.set_text_props(weight='bold')
                f_tbl.patch.set_facecolor('#f0f4f8')
                return f_tbl

            col1, col2 = st.columns([1, 1])
            
            with col1:
                st.subheader("航班起降分時預報表 (ETA/ETD)")
                fig_table1 = create_table_fig(eta_vals, etd_vals, sch_total_vals, f'{target_date_str} 航班起降分時預報表')
                st.pyplot(fig_table1)

            if has_actual:
                with col2:
                    st.subheader("航班實際起降分時表 (ATA/ATD)")
                    fig_table2 = create_table_fig(ata_vals, atd_vals, act_table_total, f'{target_date_str} 航班實際起降分時表')
                    st.pyplot(fig_table2)

            st.markdown("---")
            st.subheader("航班分時統計圖")
            
            fig_chart, ax = plt.subplots(figsize=(16, 8))
            x = np.arange(len(hours))
            width = 0.35

            ax.bar(x - width/2, eta_vals, width, label='Scheduled Arrival (預計抵達)', color='#37a794')
            ax.bar(x + width/2, etd_vals, width, label='Scheduled Departure (預計離開)', color='#32519e')
            
            ax.plot(x, sch_total_vals, marker='o', color='#7e3881', label='Total Scheduled (預計總架次)', linewidth=2)
            if has_actual:
                ax.plot(x, act_total_vals, marker='s', color='#ff7f0e', label='Total Actual (實際總架次)', linewidth=3, linestyle='--')
            
            label_fontsize = 16
            for i in range(24):
                sch_v = sch_total_vals[i]
                act_v = act_total_vals[i] if has_actual else np.nan
                
                sch_offset, sch_va = 10, 'bottom'
                act_offset, act_va = -20, 'top'
                
                if not np.isnan(act_v):
                    if act_v > sch_v:
                        act_offset, act_va = 10, 'bottom'
                        sch_offset, sch_va = -20, 'top'

                if sch_v > 0:
                    ax.annotate(str(sch_v), xy=(x[i], sch_v), xytext=(0, sch_offset), textcoords='offset points', 
                                ha='center', va=sch_va, color='#7e3881', fontweight='bold', fontsize=label_fontsize)

                if has_actual and not np.isnan(act_v) and act_v > 0:
                    ax.annotate(str(int(act_v)), xy=(x[i], act_v), xytext=(0, act_offset), textcoords='offset points', 
                                ha='center', va=act_va, color='#ff7f0e', fontweight='bold', fontsize=label_fontsize)

            ax.set_xticks(x)
            ax.set_xticklabels(labels, fontsize=11)
            ax.set_ylabel('架次', fontsize=14)
            title_text = f'{target_date_str} 航班預計與實際起降分時統計圖' if has_actual else f'{target_date_str} 航班分時預計起降統計圖'
            ax.legend(loc='upper left', bbox_to_anchor=(0, 1.05), ncol=4, frameon=False, fontsize=12)
            ax.grid(axis='y', linestyle='--', alpha=0.7)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            
            max_label_val = max(sch_total_vals)
            if has_actual:
                max_act = max([val for val in act_total_vals if not np.isnan(val)], default=0)
                max_label_val = max(max_label_val, max_act)
            ax.set_ylim(bottom=0, top=max_label_val * 1.15)
            ax.set_xlim(-0.5, len(hours) - 0.5)
            
            plt.title(title_text, fontsize=20, weight='bold', color='#142952', pad=25)
            plt.tight_layout()
            
            st.pyplot(fig_chart)
            
        except Exception as e:
            st.error(f"❌ 檔案處理失敗：{e}")
