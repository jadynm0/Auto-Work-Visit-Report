import streamlit as st
import pandas as pd
import plotly.express as px
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

# HK 18 Districts Benchmark
HK_18_DISTRICTS = [
    '中西區', '東區', '南區', '灣仔區', '九龍城', '觀塘', '深水埗', '黃大仙', '油尖旺',
    '離島', '葵青', '北區', '西貢', '沙田', '大埔', '荃灣', '屯門', '元朗'
]

def normalize_district(d_str):
    d_str = str(d_str).strip()
    mapping = {
        '灣仔': '灣仔區', '灣仔區': '灣仔區', '中環': '中西區', '銅鑼灣': '灣仔區',
        '旺角': '油尖旺', '油麻地': '油尖旺', '尖沙咀': '油尖旺', '荔枝角': '深水埗',
        '美孚': '深水埗', '大窩口': '葵青', '葵芳': '葵青', '九龍灣': '觀塘',
        '天水圍': '元朗', '黃竹坑': '南區'
    }
    return mapping.get(d_str, d_str)

if uploaded_file:
    st.success("Survey file uploaded successfully!")
    
    try:
        # ---------------------------------------------------------
        # 1. READ RAW SURVEY DATA
        # ---------------------------------------------------------
        df_raw = pd.read_excel(uploaded_file, sheet_name='表格回應 1')
        df_raw.columns = [str(c).strip() for c in df_raw.iloc[0]]
        df = df_raw.iloc[1:].reset_index(drop=True)
        
        df['店鋪_clean'] = df['店鋪'].astype(str).str.strip()
        df['姓名_clean'] = df['姓名'].astype(str).str.strip()
        df['地區_clean'] = df['地區'].apply(normalize_district)

        # SKU Column Mapping
        sku_cols = [c for c in df.columns if '架上情況 [' in str(c)]
        
        def clean_sku_name(col_name):
            match = re.search(r'\[(.*?)\]', str(col_name))
            return match.group(1).strip() if match else str(col_name).strip()

        sku_mapping = {col: clean_sku_name(col) for col in sku_cols}

        # ---------------------------------------------------------
        # 2. READ PRODUCT LISTING (DEAL STATUS) IF PROVIDED
        # ---------------------------------------------------------
        product_listing_deals = {}
        if uploaded_pl:
            try:
                df_pl = pd.read_excel(uploaded_pl, sheet_name=0)
                # Parse clients and SKUs from product listing sheet
                clients = [str(c).strip() for c in df_pl.iloc[3, 2:].values if pd.notnull(c)]
                for r_idx in range(4, len(df_pl)):
                    sku_name = str(df_pl.iloc[r_idx, 1]).strip()
                    if sku_name and sku_name != 'nan':
                        for c_idx, client in enumerate(clients):
                            has_deal = str(df_pl.iloc[r_idx, c_idx + 2]).strip().upper() == 'P'
                            if client not in product_listing_deals:
                                product_listing_deals[client] = {}
                            product_listing_deals[client][sku_name] = has_deal
            except Exception as e:
                st.info("Product listing uploaded but dynamic structure varied. Using survey defaults.")

        # ---------------------------------------------------------
        # 3. READ TARGETS DYNAMICALLY
        # ---------------------------------------------------------
        targets_dict = {}
        total_shops_dict = {}
        try:
            df_targets = pd.read_excel(uploaded_file, sheet_name='Targets')
            df_targets.columns = [str(c).strip() for c in df_targets.columns]
            
            ch_col, tg_col = df_targets.columns[0], df_targets.columns[1]
            hk_col = df_targets.columns[2] if len(df_targets.columns) > 2 else None
            
            for _, row in df_targets.dropna(subset=[ch_col]).iterrows():
                ch_raw = str(row[ch_col]).strip()
                ch_name = '7-11' if any(k in ch_raw for k in ['07-11', '7-11', '7/11', '2026-07-11']) else ch_raw
                targets_dict[ch_name] = int(float(row[tg_col])) if pd.notnull(row[tg_col]) else 0
                if hk_col and pd.notnull(row[hk_col]):
                    total_shops_dict[ch_name] = int(float(row[hk_col]))
        except Exception:
            pass

        # ---------------------------------------------------------
        # 4. DISTRICT COVERAGE ANALYSIS (X/18 DISTRICTS)
        # ---------------------------------------------------------
        visited_districts = [d for d in df['地區_clean'].unique() if d in HK_18_DISTRICTS]
        district_count = len(visited_districts)

        # ---------------------------------------------------------
        # 5. SALESPERSON BREAKDOWN & SUMMARY TABLE
        # ---------------------------------------------------------
        salespeople = [s for s in df['姓名_clean'].unique() if s and s != 'nan']
        raw_channels = df['店鋪_clean'].unique().tolist()
        
        summary_rows = []
        for ch in raw_channels:
            ch_label = '7-11' if any(k in str(ch) for k in ['7-11', '7/11', '07-11']) else ch
            if any(r['Channel'] == ch_label for r in summary_rows):
                continue
                
            sub_df = df[df['店鋪_clean'].str.contains(ch_label, na=False)]
            actual_visits = len(sub_df)
            tg_val = targets_dict.get(ch_label, 0)
            tot_shops = total_shops_dict.get(ch_label, 0)

            row_data = {
                "Channel": ch_label,
                "Total Shop in HK": tot_shops,
                "Target Visit": tg_val,
                "Actual Visit": actual_visits,
                "Status": "Target Met 🟢" if actual_visits >= tg_val and tg_val > 0 else "MISSED TARGET ❌"
            }

            # Add counts per salesperson column
            for sp in salespeople:
                sp_count = (sub_df['姓名_clean'] == sp).sum()
                row_data[sp] = sp_count

            summary_rows.append(row_data)

        df_summary = pd.DataFrame(summary_rows)

        # ---------------------------------------------------------
        # DASHBOARD SECTION 1: OVERALL KPIS & DISTRICT TRACKER
        # ---------------------------------------------------------
        st.divider()
        st.header("📌 Overall Market Visit Summary")
        
        m_col1, m_col2, m_col3 = st.columns(3)
        m_col1.metric("Total Stores Audited", len(df))
        m_col2.metric("HK District Coverage", f"{district_count}/18 Districts")
        m_col3.metric("Audited Districts", ", ".join(visited_districts[:6]) + ("..." if len(visited_districts) > 6 else ""))

        col1, col2 = st.columns([1.4, 1])
        with col1:
            st.subheader("Channel & Salesperson Performance")
            st.dataframe(df_summary, use_container_width=True)

        with col2:
            st.subheader("channel/actual visit Share")
            chart_df = df_summary[df_summary['Actual Visit'] > 0]
            fig = px.pie(chart_df, values='Actual Visit', names='Channel', hole=0.5, color_discrete_sequence=px.colors.qualitative.Set2)
            fig.update_traces(textinfo='percent+label', textposition='outside')
            st.plotly_chart(fig, use_container_width=True)

        # ---------------------------------------------------------
        # DASHBOARD SECTION 2: INTERACTIVE CHANNEL / OVERALL ANALYSIS
        # ---------------------------------------------------------
        st.divider()
        st.header("🔍 Interactive Analysis By Channel")
        
        channel_options = ["All Stores (Overall)"] + df_summary['Channel'].tolist()
        selected_ch = st.selectbox("Select Channel / View:", options=channel_options)

        if selected_ch == "All Stores (Overall)":
            df_ch = df
        else:
            df_ch = df[df['店鋪_clean'].str.contains(selected_ch, na=False)]

        tot_v = len(df_ch)

        # CHOICE COVERAGE RANGE MATRIX
        st.subheader(f"📊 {selected_ch} - Choice Coverage Breakdown")
        if tot_v > 0 and len(sku_cols) > 0:
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

            # DETAILED SKU SHELF STATUS TABLE
            st.subheader(f"🛒 {selected_ch} - SKU On-Shelf Availability")
            sku_details = []
            for orig_col, clean_name in sku_mapping.items():
                col_s = df_ch[orig_col].astype(str)
                has_stock = col_s.str.contains('有貨').sum()
                oos_tag = col_s.str.contains('缺貨').sum()
                no_tag = col_s.str.contains('無貨').sum()
                cov = round((has_stock / tot_v) * 100, 1)
                
                # Check deal status from product listing
                has_deal = product_listing_deals.get(selected_ch, {}).get(clean_name, None)
                deal_note = "With Deal (有deal)" if has_deal is True else ("No Deal" if has_deal is False else "N/A")
                if cov == 0 and has_deal is False:
                    deal_note = "0% due to No Deal (未上架/無Deal)"

                sku_details.append({
                    "Product SKU": clean_name,
                    "Deal Status": deal_note,
                    "有貨有牌仔": has_stock,
                    "缺貨有牌仔": oos_tag,
                    "無貨無牌仔": no_tag,
                    "Coverage (%)": f"{cov}%"
                })
            st.dataframe(pd.DataFrame(sku_details), use_container_width=True)

        # ---------------------------------------------------------
        # SECTION 3: MULTI-TAB BEAUTIFIED EXCEL EXPORT
        # ---------------------------------------------------------
        st.divider()
        st.header("📥 Download Complete Formatted Excel Report")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'border': 1, 'align': 'center'})
            cell_fmt = workbook.add_format({'border': 1, 'align': 'center'})

            # 1. Summary Sheet
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            ws_sum = writer.sheets['Summary']
            for c_idx, col in enumerate(df_summary.columns):
                max_len = max(df_summary[col].astype(str).map(len).max(), len(str(col))) + 4
                ws_sum.set_column(c_idx, c_idx, max(max_len, 12), cell_fmt)
                ws_sum.write(0, c_idx, col, header_fmt)

            # 2. District Breakdown Sheet
            df_dist = pd.DataFrame([{"Audited District": d, "Status": "Visited 🟢"} for d in visited_districts])
            df_dist.to_excel(writer, sheet_name='District Coverage', index=False)

            # Helper function for SKU tabs
            def export_detailed_sheet(sub_df, sheet_name):
                tot_visits = len(sub_df)
                if tot_visits == 0:
                    return

                sku_records = []
                for orig_col, clean_name in sku_mapping.items():
                    col_series = sub_df[orig_col].astype(str)
                    has_stock = col_series.str.contains('有貨').sum()
                    oos_tag = col_series.str.contains('缺貨').sum()
                    no_tag = col_series.str.contains('無貨').sum()
                    cov = round((has_stock / tot_visits) * 100, 1)
                    
                    has_deal = product_listing_deals.get(sheet_name, {}).get(clean_name, None)
                    deal_note = "With Deal" if has_deal is True else ("No Deal" if has_deal is False else "-")
                    
                    sku_records.append({
                        "Product SKU": clean_name,
                        "Deal Status": deal_note,
                        "有貨有牌仔": has_stock,
                        "缺貨有牌仔": oos_tag,
                        "無貨無牌仔": no_tag,
                        "Coverage (%)": f"{cov}%"
                    })
                
                df_out = pd.DataFrame(sku_records)
                df_out.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
                for c_idx, col in enumerate(df_out.columns):
                    max_len = max(df_out[col].astype(str).map(len).max(), len(str(col))) + 4
                    ws.set_column(c_idx, c_idx, max(max_len, 12), cell_fmt)
                    ws.write(0, c_idx, col, header_fmt)

            # Export Overall & Channels
            export_detailed_sheet(df, 'All Stores (Overall)')
            for ch_label in df_summary['Channel']:
                sub_df = df[df['店鋪_clean'].str.contains(ch_label, na=False)]
                export_detailed_sheet(sub_df, str(ch_label)[:30])

        st.download_button(
            label="🟢 Download Multi-Tab Excel Report (.xlsx)",
            data=output.getvalue(),
            file_name=f"Market_Visit_Summary_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")