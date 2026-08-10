import streamlit as st
import pandas as pd
import plotly.express as px
import xlsxwriter
import re
import io

st.set_page_config(page_title="Market Visit Summary Dashboard", page_icon="📊", layout="wide")
st.title("📊 Monthly Market Visit Performance Dashboard")

# Dual File Uploaders
col_up1, col_up2 = st.columns(2)
with col_up1:
    uploaded_file = st.file_uploader("1. Upload Market Visit Survey File (.xlsx)", type=["xlsx"])
with col_up2:
    uploaded_pl = st.file_uploader("2. Upload Product Listing File (.xlsx) [Optional]", type=["xlsx"])

# Official HK 18 Administrative Districts
HK_18_DISTRICTS = [
    '中西區', '東區', '南區', '灣仔區', '九龍城', '觀塘', '深水埗', '黃大仙', '油尖旺',
    '離島', '葵青', '北區', '西貢', '沙田', '大埔', '荃灣', '屯門', '元朗'
]

# Client Name Mapping for Bilingual Channel Names
CLIENT_NAME_MAP = {
    'Wellcome': ['Wellcome', '惠康'],
    'ParkNshop': ['ParkNshop', 'ParknShop', '百佳'],
    '7-11': ['7-11', '7/11', '07-11', '2026-07-11'],
    'Circle K': ['Circle K', 'OK'],
    '佳寶': ['佳寶'],
    'Aeon': ['Aeon', 'AEON'],
    "city'super": ["city'super", "City Super", "City\nSuper", "CitySuper"]
}

def normalize_district(d_str):
    d_str = str(d_str).strip()
    mapping = {
        '灣仔': '灣仔區', '灣仔區': '灣仔區', '中環': '中西區', '銅鑼灣': '灣仔區',
        '旺角': '油尖旺', '油麻地': '油尖旺', '尖沙咀': '油尖旺', '荔枝角': '深水埗',
        '美孚': '深水埗', '大窩口': '葵青', '葵芳': '葵青', '九龍灣': '觀塘',
        '天水圍': '元朗', '黃竹坑': '南區'
    }
    return mapping.get(d_str, d_str)

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
    
    for pl_sku, has_p in client_deals_dict.items():
        clean_p = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', str(pl_sku)).lower()
        if clean_s == clean_p or clean_s in clean_p or clean_p in clean_s:
            return has_p
    return None

if uploaded_file is not None:
    st.toast("File uploaded successfully! Processing summary...", icon="✅")
    
    try:
        # ---------------------------------------------------------
        # 1. READ RAW SURVEY DATA
        # ---------------------------------------------------------
        df_raw = pd.read_excel(uploaded_file, sheet_name='表格回應 1')
        df_raw.columns = [str(c).strip() for c in df_raw.iloc[0]]
        df = df_raw.iloc[1:].reset_index(drop=True)
        
        df['店鋪_clean'] = df['店鋪'].astype(str).str.strip()
        df['姓名_clean'] = df['姓名'].apply(normalize_salesperson)
        df['地區_clean'] = df['地區'].apply(normalize_district)

        # Extract ALL SKUs dynamically
        sku_cols = [c for c in df.columns if '架上情況 [' in str(c)]
        
        def clean_sku_name(col_name):
            match = re.search(r'\[(.*?)\]', str(col_name))
            return match.group(1).strip() if match else str(col_name).strip()

        sku_mapping = {col: clean_sku_name(col) for col in sku_cols}

        # ---------------------------------------------------------
        # 2. READ PRODUCT LISTING DEALS DYNAMICALLY
        # ---------------------------------------------------------
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
                                has_p = (val == 'P')
                                
                                std_channel = raw_client
                                for std_ch, aliases in CLIENT_NAME_MAP.items():
                                    if any(a.lower() in raw_client.lower() or raw_client.lower() in a.lower() for a in aliases):
                                        std_channel = std_ch
                                        break
                                
                                if std_channel not in product_listing_deals:
                                    product_listing_deals[std_channel] = {}
                                product_listing_deals[std_channel][sku_name] = has_p
            except Exception:
                st.info("Product listing structure varied; continuing with survey defaults.")

        # ---------------------------------------------------------
        # 3. READ TARGETS & TOTAL SHOPS DYNAMICALLY FROM FILE
        # ---------------------------------------------------------
        targets_dict = {}
        total_shops_dict = {}
        
        # Check if user file has 'summary' tab or 'Targets' sheet for total shop counts
        try:
            excel_obj = pd.ExcelFile(uploaded_file)
            if 'summary' in excel_obj.sheet_names:
                df_sum_file = pd.read_excel(uploaded_file, sheet_name='summary')
                df_sum_file.columns = [str(c).strip().lower() for c in df_sum_file.columns]
                ch_c = [c for c in df_sum_file.columns if 'channel' in c][0]
                tot_c = [c for c in df_sum_file.columns if 'total shop' in c or 'hk' in c][0]
                tg_c = [c for c in df_sum_file.columns if 'target' in c][0]
                
                for _, r in df_sum_file.dropna(subset=[ch_c]).iterrows():
                    c_name = '7-11' if any(k in str(r[ch_c]) for k in ['07-11', '7-11', '7/11', '2026-07-11']) else str(r[ch_c]).strip()
                    if pd.notnull(r[tot_c]):
                        try:
                            total_shops_dict[c_name] = int(float(r[tot_c]))
                        except:
                            pass
                    if pd.notnull(r[tg_c]):
                        try:
                            targets_dict[c_name] = int(float(r[tg_c]))
                        except:
                            pass
            elif 'Targets' in excel_obj.sheet_names:
                df_targets = pd.read_excel(uploaded_file, sheet_name='Targets')
                df_targets.columns = [str(c).strip() for c in df_targets.columns]
                ch_col, tg_col = df_targets.columns[0], df_targets.columns[1]
                hk_col = df_targets.columns[2] if len(df_targets.columns) > 2 else None
                
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
        except Exception:
            pass

        # ---------------------------------------------------------
        # 4. DISTRICT COVERAGE TRACKER
        # ---------------------------------------------------------
        visited_districts = [d for d in df['地區_clean'].unique() if d in HK_18_DISTRICTS]
        district_count = len(visited_districts)

        # ---------------------------------------------------------
        # 5. DYNAMIC SALESPERSON & CHANNEL BREAKDOWN
        # ---------------------------------------------------------
        salespeople = [s for s in df['姓名_clean'].unique() if s and s != 'nan']
        raw_channels = df['店鋪_clean'].unique().tolist()
        
        consolidated_counts = {}
        for ch in raw_channels:
            label = '7-11' if any(k in str(ch) for k in ['7-11', '7/11', '07-11']) else ch
            count = (df['店鋪_clean'] == ch).sum()
            consolidated_counts[label] = consolidated_counts.get(label, 0) + count

        summary_rows = []
        for ch_name, actual_count in consolidated_counts.items():
            tg_val = targets_dict.get(ch_name, 0)
            tot_shops = total_shops_dict.get(ch_name, "N/A")
            status = "Target Met 🟢" if actual_count >= tg_val and tg_val > 0 else ("No Target Set ⚪" if tg_val == 0 else "MISSED TARGET ❌")
            
            sub_df = df[df['店鋪_clean'].str.contains(ch_name, regex=False, na=False)]
            
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

        # Add Bottom Total Row
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

        # ---------------------------------------------------------
        # SECTION 1: OVERALL KPI DASHBOARD
        # ---------------------------------------------------------
        st.divider()
        st.header("📌 Overall Market Visit Summary")
        
        m_col1, m_col2 = st.columns([1, 2])
        m_col1.metric("Total Stores Audited", len(df))
        m_col2.metric("HK District Coverage", f"{district_count}/18 Administrative Districts Visited")

        with st.expander(f"📍 View All {district_count} Visited Districts in Hong Kong", expanded=True):
            district_tags = " • ".join([f"**{d}** 🟢" for d in visited_districts])
            st.markdown(district_tags)

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

        # ---------------------------------------------------------
        # SECTION 2: INTERACTIVE "BY CHANNEL" OR "ALL STORES (OVERALL)"
        # ---------------------------------------------------------
        st.divider()
        st.header("🔍 Interactive Analysis (Channel or Overall)")
        
        channel_options = ["All Stores (Overall)"] + [c for c in df_summary['Channel'].tolist() if c != 'Total']
        selected_ch = st.selectbox("Select View:", options=channel_options)

        if selected_ch == "All Stores (Overall)":
            df_ch = df
            st.subheader("🌐 Overall (All Brands / Shops Combined)")
            st.metric("Total Stores Audited", len(df_ch))
        else:
            df_ch = df[df['店鋪_clean'].str.contains(selected_ch, regex=False, na=False)]
            ch_target = targets_dict.get(selected_ch, 0)
            ch_actual = len(df_ch)

            m1, m2 = st.columns(2)
            m1.metric("Target Visit", ch_target)
            m2.metric("Actual Visit", ch_actual)

        # CHOICE COVERAGE RANGE MATRIX
        st.subheader(f"📊 {selected_ch} - Choice Coverage Breakdown")
        if len(df_ch) > 0 and len(sku_cols) > 0:
            in_stock_counts = df_ch[sku_cols].apply(lambda row: row.astype(str).str.contains('有貨').sum(), axis=1)
            
            c_0 = (in_stock_counts == 0).sum()
            c_1_4 = ((in_stock_counts >= 1) & (in_stock_counts <= 4)).sum()
            c_5_9 = ((in_stock_counts >= 5) & (in_stock_counts <= 9)).sum()
            c_gt_9 = (in_stock_counts > 9).sum()
            tot_v = len(df_ch)
            
            df_choice = pd.DataFrame([
                {"Choice Coverage Range": "0 SKUs", "Store Count": c_0, "Share (%)": f"{round((c_0/tot_v)*100, 1)}%"},
                {"Choice Coverage Range": "1-4 SKUs", "Store Count": c_1_4, "Share (%)": f"{round((c_1_4/tot_v)*100, 1)}%"},
                {"Choice Coverage Range": "5-9 SKUs", "Store Count": c_5_9, "Share (%)": f"{round((c_5_9/tot_v)*100, 1)}%"},
                {"Choice Coverage Range": ">9 SKUs", "Store Count": c_gt_9, "Share (%)": f"{round((c_gt_9/tot_v)*100, 1)}%"},
            ])
            st.dataframe(df_choice, use_container_width=True)

            # DETAILED SKU SHELF STATUS TABLE WITH STRICT DEAL LOGIC
            st.subheader(f"🛒 {selected_ch} - Detailed SKU Shelf Status & Performance")
            sku_details = []
            
            client_deals = product_listing_deals.get(selected_ch, {})

            for orig_col, clean_name in sku_mapping.items():
                col_s = df_ch[orig_col].astype(str)
                
                has_stock = col_s.str.contains('有貨').sum()
                oos_tag = col_s.str.contains('缺貨').sum()
                no_tag = col_s.str.contains('無貨').sum()
                cov = round((has_stock / tot_v) * 100, 1)
                
                # Strict Deal Logic
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
            st.dataframe(pd.DataFrame(sku_details), use_container_width=True)
        else:
            st.info("No visit records found.")

        # ---------------------------------------------------------
        # SECTION 3: MULTI-TAB EXCEL EXPORT WORKBOOK GENERATOR
        # ---------------------------------------------------------
        st.divider()
        st.header("📥 Download Complete Formatted Excel Summary")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'border': 1, 'align': 'center'})
            cell_fmt = workbook.add_format({'border': 1, 'align': 'center'})
            
            # 1. Summary Sheet
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            ws_summary = writer.sheets['Summary']
            
            for col_idx, col in enumerate(df_summary.columns):
                max_len = max(df_summary[col].astype(str).map(len).max(), len(str(col))) + 5
                ws_summary.set_column(col_idx, col_idx, max(max_len, 12), cell_fmt)
                ws_summary.write(0, col_idx, col, header_fmt)

            chart_col_idx = len(df_summary.columns) + 1
            chart_col_letter = xlsxwriter.utility.xl_col_to_name(chart_col_idx)
            chart_cell = f"{chart_col_letter}2"

            chart = workbook.add_chart({'type': 'doughnut'})
            max_row = len(df_summary)
            chart.add_series({
                'name':       'channel/actual visit',
                'categories': ['Summary', 1, 0, max_row - 1, 0],
                'values':     ['Summary', 1, 3, max_row - 1, 3],
                'data_labels': {'percentage': True},
            })
            chart.set_title({'name': 'channel/actual visit'})
            ws_summary.insert_chart(chart_cell, chart)

            # 2. District Breakdown Sheet
            df_dist = pd.DataFrame([{"Audited District": d, "Status": "Visited 🟢"} for d in visited_districts])
            df_dist.to_excel(writer, sheet_name='District Coverage', index=False)

            # Helper function for SKU tabs
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
                    max_len = max(
                        df_sku_details[col].astype(str).map(len).max(),
                        len(str(col))
                    ) + 5
                    ws.set_column(col_idx, col_idx, max(max_len, 15), cell_fmt)

            # Export Overall & Channel Sheets
            export_detailed_channel_sheet(df, 'All Stores (Overall)')

            for ch_label in [c for c in df_summary['Channel'] if c != 'Total']:
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