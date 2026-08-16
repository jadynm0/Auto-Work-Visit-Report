import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import xlsxwriter
import re
import io

st.set_page_config(page_title="Market Visit Summary Dashboard", page_icon="📊", layout="wide")
st.title("📊 Monthly Market Visit Performance & AI Strategic Dashboard")

# Dual File Uploaders
col_up1, col_up2 = st.columns(2)
with col_up1:
    uploaded_file = st.file_uploader("1. Upload Market Visit Survey File (.xlsx)", type=["xlsx"])
with col_up2:
    uploaded_pl = st.file_uploader("2. Upload Product Listing File (.xlsx) [Optional]", type=["xlsx"])

# 1. 18 HK Administrative Districts Grouped by Region
HK_REGION_GROUPS = {
    'Hong Kong Island (港島)': ['中西區', '東區', '南區', '灣仔區'],
    'Kowloon (九龍)': ['九龍城區', '觀塘區', '深水埗區', '黃大仙區', '油尖旺區'],
    'New Territories & Islands (新界及離島)': ['離島區', '葵青區', '北區', '西貢區', '沙田區', '大埔區', '荃灣區', '屯門區', '元朗區']
}

# 2. Standard Client Name Map for Bilingual Alignment
CLIENT_NAME_MAP = {
    'Wellcome': ['Wellcome', '惠康'],
    'ParkNshop': ['ParkNshop', 'ParknShop', '百佳'],
    '7-11': ['7-11', '7/11', '07-11', '2026-07-11'],
    'Circle K': ['Circle K', 'OK'],
    '佳寶': ['佳寶'],
    'Aeon': ['Aeon', 'AEON'],
    "city'super": ["city'super", "City Super", "City\nSuper", "CitySuper"],
    'CitiStore': ['CitiStore', '千色', 'CITI STORE'],
    '永安': ['永安', 'Wing On'],
    'UNY': ['UNY']
}

# 3. Product SKU Name Cross-Reference (Short Survey Names <-> Official Catalog Names)
SKU_NAME_MAP = {
    '310ml楊枝甘露(常溫)': ['常溫楊枝甘露310ml', '芒果椰果柚子甘露310ml'],
    '蘆楊': ['蘆薈楊枝甘露', '蘆薈楊枝甘露450ml'],
    '紫米露': ['椰香紫米甘露', '椰香紫米甘露 330g'],
    '爆檸蜜': ['凍檸蜜'],
    '花膠冰糖雪耳甘露飲品': ['花膠雪耳冰糖甘露', '花膠雪耳冰糖甘露330g'],
    '杏仁露': ['杏仁露 330ml']
}

def normalize_district_18(d_str):
    d = str(d_str).strip()
    mapping = {
        '灣仔': '灣仔區', '灣仔區': '灣仔區', '中環': '中西區', '中西區': '中西區',
        '東區': '東區', '南區': '南區', '銅鑼灣': '灣仔區',
        '旺角': '油尖旺區', '油麻地': '油尖旺區', '尖沙咀': '油尖旺區', '油尖旺': '油尖旺區', '油尖旺區': '油尖旺區',
        '荔枝角': '深水埗區', '美孚': '深水埗區', '深水埗': '深水埗區', '深水埗區': '深水埗區',
        '九龍城': '九龍城區', '九龍城區': '九龍城區',
        '黃大仙': '黃大仙區', '黃大仙區': '黃大仙區',
        '九龍灣': '觀塘區', '觀塘': '觀塘區', '觀塘區': '觀塘區',
        '大窩口': '葵青區', '葵芳': '葵青區', '葵青': '葵青區', '葵青區': '葵青區',
        '天水圍': '元朗區', '元朗': '元朗區', '元朗區': '元朗區',
        '黃竹坑': '南區', '屯門': '屯門區', '屯門區': '屯門區',
        '荃灣': '荃灣區', '荃灣區': '荃灣區', '大埔': '大埔區', '大埔區': '大埔區',
        '沙田': '沙田區', '沙田區': '沙田區', '北區': '北區', '西貢': '西貢區', '西貢區': '西貢區', '離島': '離島區'
    }
    return mapping.get(d, d)

def normalize_salesperson(sp_str):
    s = str(sp_str).strip()
    if s.lower() in ['chris wong', 'chriswong']:
        return 'Chris Wong'
    if s.lower() in ['eunice chu', 'eunice', 'eunicfg']:
        return 'Eunice Chu'
    return s

def get_coverage_rating(cov_num):
    if cov_num >= 85:
        return "Good 🟢"
    elif cov_num >= 80:
        return "Satisfy 🔵"
    elif cov_num >= 70:
        return "OK 🟡"
    else:
        return "Below Benchmark 🔴"

def match_sku_deal(survey_sku, client_deals_dict):
    if not client_deals_dict:
        return None
    clean_s = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', str(survey_sku)).lower()
    aliases = SKU_NAME_MAP.get(survey_sku, [survey_sku])
    for pl_sku, has_p in client_deals_dict.items():
        clean_p = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', str(pl_sku)).lower()
        if any(re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', a).lower() in clean_p for a in aliases):
            return has_p
        if clean_s == clean_p or clean_s in clean_p or clean_p in clean_s:
            return has_p
    return None

def parse_targets_dynamically(df_targets):
    """Dynamically parses targets and total stores regardless of header row offsets."""
    targets_dict = {}
    total_shops_dict = {}
    
    col_str = " ".join([str(c).lower() for c in df_targets.columns])
    if 'channel' in col_str or 'target' in col_str or '渠道' in col_str or '目標' in col_str:
        header_row_idx = None
    else:
        header_row_idx = None
        for r_idx in range(min(5, len(df_targets))):
            row_vals = [str(v).lower() for v in df_targets.iloc[r_idx].values if pd.notnull(v)]
            if any('channel' in v or '店' in v or '渠道' in v for v in row_vals) and any('target' in v or '目標' in v for v in row_vals):
                header_row_idx = r_idx
                break
                
    if header_row_idx is not None:
        df_targets.columns = [str(c).strip() for c in df_targets.iloc[header_row_idx]]
        df_targets = df_targets.iloc[header_row_idx+1:].reset_index(drop=True)
    else:
        df_targets.columns = [str(c).strip() for c in df_targets.columns]
        
    ch_col = None
    tg_col = None
    hk_col = None
    
    for c in df_targets.columns:
        c_low = str(c).lower()
        if 'channel' in c_low or '店' in c_low or '渠道' in c_low:
            ch_col = c
        elif 'target' in c_low or '目標' in c_low:
            if tg_col is None:
                tg_col = c
        elif 'total' in c_low or 'hk' in c_low or '總數' in c_low:
            hk_col = c
            
    if ch_col is None and len(df_targets.columns) > 0:
        ch_col = df_targets.columns[0]
    if tg_col is None and len(df_targets.columns) > 1:
        tg_col = df_targets.columns[1]
        
    for _, row in df_targets.dropna(subset=[ch_col]).iterrows():
        ch_raw = str(row[ch_col]).strip()
        ch_name = '7-11' if any(k in ch_raw for k in ['07-11', '7-11', '7/11', '2026-07-11']) else ch_raw
        try:
            targets_dict[ch_name] = int(float(row[tg_col]))
        except:
            targets_dict[ch_name] = 0
            
        if hk_col and pd.notnull(row[hk_col]):
            try:
                total_shops_dict[ch_name] = int(float(row[hk_col]))
            except:
                pass
                
    return targets_dict, total_shops_dict

def generate_dynamic_market_insights(df, df_summary, df_dist_summary, sku_mapping):
    """Pure dynamic FMCG analytics synthesizer with zero cloud API dependencies."""
    total_audits = len(df)
    visited_dist = len(df_dist_summary[df_dist_summary['Status'] == 'Visited 🟢'])
    
    missed_channels = df_summary[(df_summary['Channel'] != 'Total') & (df_summary['Status'].str.contains('MISSED|Unvisited'))]['Channel'].tolist()
    met_channels = df_summary[(df_summary['Channel'] != 'Total') & (df_summary['Status'].str.contains('Target Met'))]['Channel'].tolist()
    
    sku_stats = []
    for orig_col, clean_name in sku_mapping.items():
        col_s = df[orig_col].astype(str)
        in_stk = col_s.str.contains('有貨').sum()
        oos = col_s.str.contains('缺貨').sum()
        cov = round((in_stk / total_audits) * 100, 1) if total_audits > 0 else 0
        sku_stats.append((clean_name, in_stk, oos, cov))
        
    top_skus = sorted(sku_stats, key=lambda x: x[3], reverse=True)[:5]
    oos_skus = sorted([s for s in sku_stats if s[2] > 0], key=lambda x: x[2], reverse=True)[:5]
    zero_cov_skus = [s for s in sku_stats if s[3] == 0]
    
    ch_insights = []
    for ch in ['7-11', 'Circle K', 'Wellcome', 'ParkNshop', "city'super", 'Aeon']:
        sub = df[df['店鋪_clean'].str.contains(ch, regex=False, na=False)]
        if len(sub) > 0:
            top_ch_sku = []
            for orig_col, clean_name in sku_mapping.items():
                s_cnt = sub[orig_col].astype(str).str.contains('有貨').sum()
                rate = round((s_cnt / len(sub)) * 100, 1)
                top_ch_sku.append((clean_name, rate))
            top_ch_sku.sort(key=lambda x: x[1], reverse=True)
            zeros = [x for x in top_ch_sku if x[1] == 0]
            ch_insights.append((ch, len(sub), top_ch_sku[0], top_ch_sku[1], len(zeros)))

    text = f"""### 📊 一、 市場分析 (Market Analysis & Findings)
* **通路覆蓋與執行率**:
  * 全港實地走訪共 **{total_audits} 間分店**，覆蓋全港 **{visited_dist}/18 個行政區**。
  * **目標達成情況**: {('、'.join(met_channels) + ' 順利達成預定指標') if met_channels else '各主要通路走訪數量仍有提升空間'}；而 **{('、'.join(missed_channels))}** 走訪進度落後或尚未有走訪紀錄，反映外勤資源需進一步平衡。
* **各主要零售通路表現**:
"""
    for ch, count, top1, top2, z_count in ch_insights:
        text += f"  * **{ch}** ({count} 間分店): 主力熱銷為 **{top1[0]}** ({top1[1]}% 覆蓋) 及 **{top2[0]}** ({top2[1]}%)；另有 {z_count} 款產品未見陳列 (0%)。\n"

    text += f"""* **主力品項與斷貨預警**:
  * **全港覆蓋率最高 SKU**: {', '.join([f'{s[0]} ({s[3]}%)' for s in top_skus])}。
  * **缺貨有牌仔 (OOS) 警示**: {', '.join([f'{s[0]} ({s[2]}間分店缺貨)' for s in oos_skus]) if oos_skus else '各通路供貨穩定，未見明顯標籤缺貨'}。

---

### 💡 二、 行銷與營運建議 (Strategic Recommendations)
* **1. 旺季供應鏈與安全庫存預警**:
  * 針對門市出現缺貨牌仔的熱銷品項（如 **{', '.join([s[0] for s in oos_skus[:3]]) if oos_skus else '主力涼茶系列'}**），建議提早評估產能及補貨頻率，避免夏季高峰期因門市斷貨遭競品擠佔。
* **2. 長尾品項精簡與陳列優化 (SKU Rationalization)**:
  * 目前全港有 **{len(zero_cov_skus)} 款 SKU** 覆蓋率為 0%。建議審視無合約上架 (No Deal) 與低轉速產品，將冷櫃陳列面集中配置於高週轉皇牌。
* **3. 18區外勤巡查路線動態排班**:
  * 針對未達標或未走訪通路建立跨區輪替排班，平衡港島、九龍與新界走訪密度以確保審查數據全面。
"""
    return text

if uploaded_file is not None:
    st.toast("File uploaded successfully! Processing summary...", icon="✅")
    
    try:
        # 1. READ RAW SURVEY DATA
        df_raw = pd.read_excel(uploaded_file, sheet_name='表格回應 1')
        df_raw.columns = [str(c).strip() for c in df_raw.iloc[0]]
        df = df_raw.iloc[1:].reset_index(drop=True)
        
        df['店鋪_clean'] = df['店鋪'].astype(str).str.strip()
        df['姓名_clean'] = df['姓名'].apply(normalize_salesperson)
        df['地區_clean'] = df['地區'].apply(normalize_district_18)

        sku_cols = [c for c in df.columns if '架上情況 [' in str(c)]
        
        def clean_sku_name(col_name):
            match = re.search(r'\[(.*?)\]', str(col_name))
            return match.group(1).strip() if match else str(col_name).strip()

        sku_mapping = {col: clean_sku_name(col) for col in sku_cols}

        # 2. READ PRODUCT LISTING DEALS
        product_listing_deals = {}
        if uploaded_pl is not None:
            try:
                df_pl = pd.read_excel(uploaded_pl, sheet_name=0)
                clients_raw = [str(c).replace('\n', ' ').strip() for c in df_pl.iloc[3, 2:].values if pd.notnull(c)]
                for r_idx in range(4, len(df_pl)):
                    sku_name = str(df_pl.iloc[r_idx, 1]).strip()
                    if sku_name and sku_name != 'nan':
                        for c_idx, raw_client in enumerate(clients_raw):
                            if (c_idx + 2) < len(df_pl.columns):
                                val = str(df_pl.iloc[r_idx, c_idx + 2]).strip().upper()
                                std_channel = raw_client
                                for std_ch, aliases in CLIENT_NAME_MAP.items():
                                    if any(a.lower() in raw_client.lower() or raw_client.lower() in a.lower() for a in aliases):
                                        std_channel = std_ch
                                        break
                                if std_channel not in product_listing_deals:
                                    product_listing_deals[std_channel] = {}
                                product_listing_deals[std_channel][sku_name] = (val == 'P')
            except Exception:
                pass

        # 3. DYNAMIC TARGET & TOTAL STORE PARSER
        targets_dict = {}
        total_shops_dict = {}
        excel_obj = pd.ExcelFile(uploaded_file)
        
        target_sheets = [s for s in excel_obj.sheet_names if any(k in s.lower() for k in ['target', 'summary', '目標'])]
        if target_sheets:
            df_t_raw = pd.read_excel(uploaded_file, sheet_name=target_sheets[0])
            targets_dict, total_shops_dict = parse_targets_dynamically(df_t_raw)

        # 4. DISTRICT BREAKDOWN GROUPED BY REGION
        dist_counts_series = df['地區_clean'].value_counts()
        district_summary_rows = []
        visited_count_18 = 0

        for region_name, dist_list in HK_REGION_GROUPS.items():
            for dist in dist_list:
                v_count = int(dist_counts_series.get(dist, 0))
                is_visited = v_count > 0
                if is_visited:
                    visited_count_18 += 1
                district_summary_rows.append({
                    "Region (區域)": region_name,
                    "Administrative District (18區)": dist,
                    "Status": "Visited 🟢" if is_visited else "Unvisited ⚪",
                    "Audit Count (走訪分店數)": v_count
                })

        df_dist_summary = pd.DataFrame(district_summary_rows)

        # 5. DYNAMIC CHANNEL & SALESPERSON TABLE (INCLUDING UNVISITED TARGET SHOPS)
        salespeople = [s for s in df['姓名_clean'].unique() if s and s != 'nan']
        visited_channels = df['店鋪_clean'].unique().tolist()
        all_channels = list(dict.fromkeys(visited_channels + list(targets_dict.keys())))

        summary_rows = []
        for ch in all_channels:
            ch_name = '7-11' if any(k in str(ch) for k in ['7-11', '7/11', '07-11']) else str(ch)
            if any(r['Channel'] == ch_name for r in summary_rows):
                continue
                
            sub_df = df[df['店鋪_clean'].str.contains(ch_name, regex=False, na=False)]
            actual_count = len(sub_df)
            tg_val = targets_dict.get(ch_name, 0)
            tot_shops = total_shops_dict.get(ch_name, "N/A")
            
            if actual_count >= tg_val and tg_val > 0:
                status = "Target Met 🟢"
            elif tg_val == 0 and actual_count > 0:
                status = "No Target Set ⚪"
            elif actual_count == 0:
                status = "Unvisited / Missed ❌"
            else:
                status = "MISSED TARGET ❌"

            row_data = {
                "Channel": ch_name,
                "Total Shop in HK": tot_shops,
                "Target Visit": tg_val,
                "Actual Visit": actual_count,
                "Status": status
            }

            for sp in salespeople:
                row_data[sp] = (sub_df['姓名_clean'] == sp).sum()

            summary_rows.append(row_data)

        # Total Row
        total_row = {
            "Channel": "Total",
            "Total Shop in HK": sum(r["Total Shop in HK"] for r in summary_rows if isinstance(r["Total Shop in HK"], int)),
            "Target Visit": sum(r["Target Visit"] for r in summary_rows),
            "Actual Visit": sum(r["Actual Visit"] for r in summary_rows),
            "Status": "Total Summary"
        }
        for sp in salespeople:
            total_row[sp] = sum(r[sp] for r in summary_rows)
            
        summary_rows.append(total_row)
        df_summary = pd.DataFrame(summary_rows)

        ai_generated_text = generate_dynamic_market_insights(df, df_summary, df_dist_summary, sku_mapping)

        # ---------------------------------------------------------
        # TABS: PERFORMANCE DASHBOARD & DYNAMIC AI INSIGHTS
        # ---------------------------------------------------------
        tab_dash, tab_ai = st.tabs(["📊 Performance Dashboard", "💡 AI Views, Analysis & Recommendations"])

        with tab_dash:
            st.header("📌 Overall Market Visit Summary")
            
            m_col1, m_col2 = st.columns([1, 2])
            m_col1.metric("Total Stores Audited", len(df))
            m_col2.metric("HK 18-District Coverage", f"{visited_count_18}/18 Administrative Districts Audited")

            with st.expander("📍 View Regional & District Breakdown (港島 / 九龍 / 新界及離島)", expanded=True):
                st.dataframe(df_dist_summary, use_container_width=True)

            col1, col2 = st.columns([1.4, 1])
            with col1:
                st.subheader("Channel & Salesperson Form Breakdown")
                st.dataframe(df_summary, use_container_width=True)

            with col2:
                st.subheader("channel/actual visit Share")
                chart_df = df_summary[(df_summary['Channel'] != 'Total') & (df_summary['Actual Visit'] > 0)]
                if len(chart_df) > 0:
                    fig = px.pie(chart_df, values='Actual Visit', names='Channel', hole=0.5, color_discrete_sequence=px.colors.qualitative.Set2)
                    fig.update_traces(textinfo='percent+label', textposition='outside')
                    st.plotly_chart(fig, use_container_width=True)

            st.divider()
            st.header("🔍 Interactive Channel Analysis & Visual Stock Breakdown")
            
            channel_options = ["All Stores (Overall)"] + [c for c in df_summary['Channel'].tolist() if c != 'Total' and df_summary.loc[df_summary['Channel']==c, 'Actual Visit'].values[0] > 0]
            selected_ch = st.selectbox("Select View:", options=channel_options)

            if selected_ch == "All Stores (Overall)":
                df_ch = df
            else:
                df_ch = df[df['店鋪_clean'].str.contains(selected_ch, regex=False, na=False)]

            tot_v = len(df_ch)

            # Choice Coverage Matrix
            st.subheader(f"📊 {selected_ch} - Choice Coverage Breakdown")
            in_stock_counts = df_ch[sku_cols].apply(lambda row: row.astype(str).str.contains('有貨').sum(), axis=1)
            c_0 = (in_stock_counts == 0).sum()
            c_1_4 = ((in_stock_counts >= 1) & (in_stock_counts <= 4)).sum()
            c_5_9 = ((in_stock_counts >= 5) & (in_stock_counts <= 9)).sum()
            c_gt_9 = (in_stock_counts > 9).sum()
            
            df_choice = pd.DataFrame([
                {"Choice Coverage Range": "0 SKUs", "Store Count": c_0, "Share (%)": f"{round((c_0/tot_v)*100, 1)}%"},
                {"Choice Coverage Range": "1-4 SKUs", "Store Count": c_1_4, "Share (%)": f"{round((c_1_4/tot_v)*100, 1)}%"},
                {"Choice Coverage Range": "5-9 SKUs", "Store Count": c_5_9, "Share (%)": f"{round((c_5_9/tot_v)*100, 1)}%"},
                {"Choice Coverage Range": ">9 SKUs", "Store Count": c_gt_9, "Share (%)": f"{round((c_gt_9/tot_v)*100, 1)}%"},
            ])
            st.dataframe(df_choice, use_container_width=True)

            # SKU Shelf Table & Chart
            st.subheader(f"🛒 {selected_ch} - Detailed SKU Availability")
            sku_details = []
            client_deals = product_listing_deals.get(selected_ch, {})

            for orig_col, clean_name in sku_mapping.items():
                col_s = df_ch[orig_col].astype(str)
                has_stock = col_s.str.contains('有貨').sum()
                oos_tag = col_s.str.contains('缺貨').sum()
                no_tag = col_s.str.contains('無貨').sum()
                cov = round((has_stock / tot_v) * 100, 1) if tot_v > 0 else 0
                
                if selected_ch == "All Stores (Overall)":
                    deal_note = "Aggregated Overall"
                else:
                    has_deal = match_sku_deal(clean_name, client_deals)
                    if has_stock > 0 and has_deal is True:
                        deal_note = "With Deal"
                    elif has_stock > 0 and has_deal is False:
                        deal_note = "Unlisted / Stocked without Deal (無Deal有貨)"
                    elif has_stock == 0 and has_deal is True:
                        deal_note = "With Deal (OOS)"
                    elif has_stock == 0 and has_deal is False:
                        deal_note = "0% due to No Deal (無Deal未上架)"
                    else:
                        deal_note = "-"

                sku_details.append({
                    "Product SKU": clean_name,
                    "Deal Status": deal_note,
                    "有貨有牌仔": has_stock,
                    "缺貨有牌仔": oos_tag,
                    "無貨無牌仔": no_tag,
                    "Coverage (%)": f"{cov}%",
                    "Performance Rating": get_coverage_rating(cov)
                })
            
            df_sku_view = pd.DataFrame(sku_details)
            st.dataframe(df_sku_view, use_container_width=True)

            if selected_ch != "All Stores (Overall)":
                st.subheader(f"📈 {selected_ch} - SKU Stock Status Visual Chart")
                top_chart_df = df_sku_view.head(20)
                fig_bar = go.Figure(data=[
                    go.Bar(name='有貨有牌仔 (In-Stock)', x=top_chart_df['Product SKU'], y=top_chart_df['有貨有牌仔'], marker_color='#2ca02c'),
                    go.Bar(name='缺貨有牌仔 (OOS)', x=top_chart_df['Product SKU'], y=top_chart_df['缺貨有牌仔'], marker_color='#d62728'),
                    go.Bar(name='無貨無牌仔 (No Stock)', x=top_chart_df['Product SKU'], y=top_chart_df['無貨無牌仔'], marker_color='#7f7f7f')
                ])
                fig_bar.update_layout(barmode='stack', xaxis_tickangle=-45, height=500)
                st.plotly_chart(fig_bar, use_container_width=True)

        with tab_ai:
            st.header("💡 AI Executive Summary: Views, Analysis & Strategic Recommendations")
            st.markdown(ai_generated_text)

        # ---------------------------------------------------------
        # SECTION 3: MULTI-TAB EXCEL EXPORT WORKBOOK
        # ---------------------------------------------------------
        st.divider()
        st.header("📥 Download Complete Formatted Excel Report")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'border': 1, 'align': 'center'})
            cell_fmt = workbook.add_format({'border': 1, 'align': 'center'})
            ai_title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#1F497D'})
            ai_text_fmt = workbook.add_format({'text_wrap': True, 'font_size': 11})

            # 1. Summary Sheet
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            ws_summary = writer.sheets['Summary']
            for col_idx, col in enumerate(df_summary.columns):
                max_len = max(df_summary[col].astype(str).map(len).max(), len(str(col))) + 5
                ws_summary.set_column(col_idx, col_idx, max(max_len, 12), cell_fmt)
                ws_summary.write(0, col_idx, col, header_fmt)

            chart_col_idx = len(df_summary.columns) + 1
            chart_col_letter = xlsxwriter.utility.xl_col_to_name(chart_col_idx)
            chart = workbook.add_chart({'type': 'doughnut'})
            max_row = len(df_summary)
            chart.add_series({
                'name':       'channel/actual visit',
                'categories': ['Summary', 1, 0, max_row - 1, 0],
                'values':     ['Summary', 1, 3, max_row - 1, 3],
                'data_labels': {'percentage': True},
            })
            chart.set_title({'name': 'channel/actual visit'})
            ws_summary.insert_chart(f"{chart_col_letter}2", chart)

            # 2. Regional & 18-District Coverage Sheet
            df_dist_summary.to_excel(writer, sheet_name='District Coverage', index=False)
            ws_dist = writer.sheets['District Coverage']
            for col_idx, col in enumerate(df_dist_summary.columns):
                max_len = max(df_dist_summary[col].astype(str).map(len).max(), len(str(col))) + 5
                ws_dist.set_column(col_idx, col_idx, max(max_len, 15), cell_fmt)
                ws_dist.write(0, col_idx, col, header_fmt)

            # 3. AI Generated Insights Sheet
            ws_ai = workbook.add_worksheet('AI Views & Recommendations')
            ws_ai.set_column('A:A', 115)
            ws_ai.write('A1', '💡 Monthly Market Report: AI Views, Analysis & Strategic Recommendations', ai_title_fmt)
            
            clean_lines = [line.strip() for line in ai_generated_text.split('\n') if line.strip()]
            for r_idx, line in enumerate(clean_lines, start=3):
                if line.startswith("#") or line.startswith("📊") or line.startswith("💡") or line.startswith("* **"):
                    ws_ai.write(r_idx, 0, line.replace('#', '').strip(), workbook.add_format({'bold': True, 'font_size': 11, 'font_color': '#1F497D'}))
                else:
                    ws_ai.write(r_idx, 0, line, ai_text_fmt)

            # 4. Helper Function for Channel Tabs with Embedded Stacked Charts
            def export_detailed_channel_sheet(sub_df, sheet_name):
                tot_visits = len(sub_df)
                if tot_visits == 0:
                    return

                in_stock_counts = sub_df[sku_cols].apply(lambda row: row.astype(str).str.contains('有貨').sum(), axis=1)
                c_0 = (in_stock_counts == 0).sum()
                c_1_4 = ((in_stock_counts >= 1) & (in_stock_counts <= 4)).sum()
                c_5_9 = ((in_stock_counts >= 5) & (in_stock_counts <= 9)).sum()
                c_gt_9 = (in_stock_counts > 9).sum()

                df_choice_matrix = pd.DataFrame([
                    {"Choice Coverage Range": "0 SKUs", "Store Count": c_0, "Share (%)": f"{round((c_0/tot_visits)*100, 1)}%"},
                    {"Choice Coverage Range": "1-4 SKUs", "Store Count": c_1_4, "Share (%)": f"{round((c_1_4/tot_visits)*100, 1)}%"},
                    {"Choice Coverage Range": "5-9 SKUs", "Store Count": c_5_9, "Share (%)": f"{round((c_5_9/tot_visits)*100, 1)}%"},
                    {"Choice Coverage Range": ">9 SKUs", "Store Count": c_gt_9, "Share (%)": f"{round((c_gt_9/tot_visits)*100, 1)}%"},
                ])

                client_deals = product_listing_deals.get(sheet_name, {})
                sku_records = []
                for orig_col, clean_name in sku_mapping.items():
                    col_series = sub_df[orig_col].astype(str)
                    has_stock = col_series.str.contains('有貨').sum()
                    oos_tag = col_series.str.contains('缺貨').sum()
                    no_tag = col_series.str.contains('無貨').sum()
                    cov = round((has_stock / tot_visits) * 100, 1) if tot_visits > 0 else 0
                    
                    if sheet_name == "All Stores (Overall)":
                        deal_note = "Aggregated Overall"
                    else:
                        has_deal = match_sku_deal(clean_name, client_deals)
                        if has_stock > 0 and has_deal is True:
                            deal_note = "With Deal"
                        elif has_stock > 0 and has_deal is False:
                            deal_note = "Unlisted / Stocked without Deal (無Deal有貨)"
                        elif has_stock == 0 and has_deal is True:
                            deal_note = "With Deal (OOS)"
                        elif has_stock == 0 and has_deal is False:
                            deal_note = "0% due to No Deal (無Deal未上架)"
                        else:
                            deal_note = "-"

                    sku_records.append({
                        "Product SKU": clean_name,
                        "Deal Status": deal_note,
                        "有貨有牌仔": has_stock,
                        "缺貨有牌仔": oos_tag,
                        "無貨無牌仔": no_tag,
                        "Coverage (%)": f"{cov}%",
                        "Performance Rating": get_coverage_rating(cov)
                    })
                df_sku_details = pd.DataFrame(sku_records)

                df_choice_matrix.to_excel(writer, sheet_name=sheet_name, startrow=0, index=False)
                df_sku_details.to_excel(writer, sheet_name=sheet_name, startrow=7, index=False)

                ws = writer.sheets[sheet_name]
                for col_idx, col in enumerate(df_choice_matrix.columns):
                    ws.write(0, col_idx, col, header_fmt)
                for col_idx, col in enumerate(df_sku_details.columns):
                    ws.write(7, col_idx, col, header_fmt)
                for col_idx, col in enumerate(df_sku_details.columns):
                    max_len = max(df_sku_details[col].astype(str).map(len).max(), len(str(col))) + 5
                    ws.set_column(col_idx, col_idx, max(max_len, 15), cell_fmt)

                # Channel Visual Chart
                if sheet_name != "All Stores (Overall)":
                    chart_shop = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
                    max_sku_row = len(df_sku_details) + 7
                    chart_shop.add_series({
                        'name':       [sheet_name, 7, 2],
                        'categories': [sheet_name, 8, 0, min(max_sku_row, 28), 0],
                        'values':     [sheet_name, 8, 2, min(max_sku_row, 28), 2],
                        'fill':       {'color': '#2ca02c'}
                    })
                    chart_shop.add_series({
                        'name':       [sheet_name, 7, 3],
                        'categories': [sheet_name, 8, 0, min(max_sku_row, 28), 0],
                        'values':     [sheet_name, 8, 3, min(max_sku_row, 28), 3],
                        'fill':       {'color': '#d62728'}
                    })
                    chart_shop.add_series({
                        'name':       [sheet_name, 7, 4],
                        'categories': [sheet_name, 8, 0, min(max_sku_row, 28), 0],
                        'values':     [sheet_name, 8, 4, min(max_sku_row, 28), 4],
                        'fill':       {'color': '#a6a6a6'}
                    })
                    chart_shop.set_title({'name': f'{sheet_name} - SKU Shelf Availability (Top 20 SKUs)'})
                    chart_shop.set_size({'width': 700, 'height': 350})
                    ws.insert_chart('I8', chart_shop)

            # Export Overall & Channel Sheets
            export_detailed_channel_sheet(df, 'All Stores (Overall)')

            for ch_label in [c for c in df_summary['Channel'] if c != 'Total' and df_summary.loc[df_summary['Channel']==c, 'Actual Visit'].values[0] > 0]:
                sub_df = df[df['店鋪_clean'].str.contains(ch_label, regex=False, na=False)]
                sheet_title = str(ch_label).replace(':', '').replace('/', '-')[:30]
                export_detailed_channel_sheet(sub_df, sheet_title)

        st.download_button(
            label="🟢 Download Complete Formatted Excel Report (.xlsx)",
            data=output.getvalue(),
            file_name=f"Market_Visit_Summary_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")