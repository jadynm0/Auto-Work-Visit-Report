import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Market Visit Summary Dashboard", page_icon="📊", layout="wide")
st.title("📊 Monthly Market Visit Performance Dashboard")

uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    st.success("File uploaded successfully! Processing summary...")
    
    try:
        # ---------------------------------------------------------
        # 1. READ TARGETS & TOTAL STORES
        # ---------------------------------------------------------
        df_targets = pd.read_excel(uploaded_file, sheet_name='Targets')
        df_targets.columns = [str(c).strip() for c in df_targets.columns]
        
        ch_col = df_targets.columns[0]
        tg_col = df_targets.columns[1]
        
        # Optional total HK shops column if present in Targets sheet
        hk_shops_col = df_targets.columns[2] if len(df_targets.columns) > 2 else None

        df_targets = df_targets.dropna(subset=[ch_col, tg_col])
        
        cleaned_targets = {}
        total_hk_shops = {}
        
        for _, row in df_targets.iterrows():
            ch_raw = str(row[ch_col]).strip()
            ch_name = '7-11' if any(k in ch_raw for k in ['07-11', '7-11', '7/11', '2026-07-11']) else ch_raw
            
            cleaned_targets[ch_name] = int(float(row[tg_col])) if pd.notnull(row[tg_col]) else 0
            if hk_shops_col and pd.notnull(row[hk_shops_col]):
                total_hk_shops[ch_name] = int(float(row[hk_shops_col]))

        # Default HK Shop counts if not provided in sheet
        default_hk = {'Wellcome': 300, 'ParkNshop': 200, '7-11': 900, 'Circle K': 300, '佳寶': 100, 'Aeon': 12, "city'super": 6}
        for k, v in default_hk.items():
            if k not in total_hk_shops:
                total_hk_shops[k] = v

        # ---------------------------------------------------------
        # 2. READ SURVEY DATA
        # ---------------------------------------------------------
        df_raw = pd.read_excel(uploaded_file, sheet_name='表格回應 1')
        df_raw.columns = df_raw.iloc[0]
        df = df_raw.iloc[1:].reset_index(drop=True)
        df['店鋪_clean'] = df['店鋪'].astype(str).str.strip()
        
        store_counts = df['店鋪_clean'].value_counts()

        summary_rows = []
        for channel, target in cleaned_targets.items():
            if channel == '7-11':
                actual = sum(count for store, count in store_counts.items() if any(k in store for k in ['7-11', '7/11', '07-11']))
            elif channel == 'SMKT':
                actual = store_counts.get('Wellcome', 0) + store_counts.get('ParkNshop', 0)
            elif channel.lower() in ['min.chain', 'min chain']:
                actual = store_counts.get('Aeon', 0) + store_counts.get("city'super", 0)
            else:
                actual = store_counts.get(channel, 0)

            total_shop = total_hk_shops.get(channel, 100)
            coverage_pct = round((actual / total_shop) * 100, 1) if total_shop > 0 else 0

            summary_rows.append({
                "Channel": channel,
                "Total Shop in HK": total_shop,
                "Target Visit": target,
                "Actual Visit": actual,
                "Coverage (%)": coverage_pct,
                "Status": "Target Met 🟢" if actual >= target else "MISSED TARGET ❌"
            })

        df_summary = pd.DataFrame(summary_rows)

        # ---------------------------------------------------------
        # SECTION 1: OVERALL SUMMARY & INTERACTIVE PLOTLY DONUT CHART
        # ---------------------------------------------------------
        st.divider()
        st.header("📌 Overall Market Visit Summary")
        
        col1, col2 = st.columns([1.2, 1])
        
        with col1:
            st.subheader("Channel Visit Performance")
            st.dataframe(df_summary, use_container_width=True)
            st.metric("Total Stores Audited Across HK", len(df))

        with col2:
            st.subheader("channel/actual visit Share")
            chart_df = df_summary[df_summary['Actual Visit'] > 0]
            
            # Plotly Donut Chart (Fixes font bug completely)
            fig = px.pie(
                chart_df, 
                values='Actual Visit', 
                names='Channel', 
                hole=0.5,
                color_discrete_sequence=px.colors.qualitative.Set2
            )
            fig.update_traces(textinfo='percent+label', textposition='outside')
            fig.update_layout(showlegend=False, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig, use_container_width=True)

        # ---------------------------------------------------------
        # SECTION 2: INTERACTIVE "BY CHANNEL" ANALYSIS (GG'S ROW 47)
        # ---------------------------------------------------------
        st.divider()
        st.header("🔍 Interactive Analysis By Channel")
        
        channel_options = df_summary['Channel'].tolist()
        selected_ch = st.selectbox("By channel:", options=channel_options, index=channel_options.index('佳寶') if '佳寶' in channel_options else 0)

        # Filter data for selected channel
        ch_row = df_summary[df_summary['Channel'] == selected_ch].iloc[0]
        df_ch = df[df['店鋪_clean'].str.contains(selected_ch, na=False)] if selected_ch != 'SMKT' else df[df['店鋪_clean'].isin(['Wellcome', 'ParkNshop'])]

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Shop in HK", ch_row['Total Shop in HK'])
        m2.metric("Target Visit", ch_row['Target Visit'])
        m3.metric("Actual Visit", ch_row['Actual Visit'])
        m4.metric("探店 Coverage Rate", f"{ch_row['Coverage (%)']}%")

        # ---------------------------------------------------------
        # CHOICE COVERAGE RANGE MATRIX (0, 1-4, 5-9, >9)
        # ---------------------------------------------------------
        st.subheader(f"📊 {selected_ch} - Choice Coverage Breakdown")
        
        # Calculate SKU counts per store
        sku_cols = [c for c in df_ch.columns if str(c).startswith('架上情況 [')]
        
        if len(df_ch) > 0 and len(sku_cols) > 0:
            # Count in-stock products for each store visit
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

            # ---------------------------------------------------------
            # DETAILED SKU SHELF STATUS TABLE FOR SELECTED CHANNEL
            # ---------------------------------------------------------
            st.subheader(f"🛒 {selected_ch} - Detailed SKU Shelf Status")
            
            sku_details = []
            for col in sku_cols:
                clean_sku = col.replace('架上情況 [', '').replace(']', '').strip()
                col_s = df_ch[col].astype(str)
                
                has_stock = col_s.str.contains('有貨').sum()
                oos_tag = col_s.str.contains('缺貨').sum()
                no_tag = col_s.str.contains('無貨').sum()
                
                cov = round((has_stock / tot_v) * 100, 1)
                sku_details.append({
                    "Product SKU": clean_sku,
                    "有貨有牌仔": has_stock,
                    "缺貨有牌仔": oos_tag,
                    "無貨無牌仔 / 無貨唔牌仔": no_tag,
                    "Selling Coverage (%)": f"{cov}%"
                })
                
            st.dataframe(pd.DataFrame(sku_details), use_container_width=True)
        else:
            st.info("No visit response records found for this selected channel.")

       # ---------------------------------------------------------
        # DOWNLOAD MULTI-TAB EXCEL WORKBOOK
        # ---------------------------------------------------------
        st.divider()
        st.header("📥 Download Excel Summary")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # TAB 1: Channel Summary & Donut Chart
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            
            workbook = writer.book
            worksheet_summary = writer.sheets['Summary']
            
            chart = workbook.add_chart({'type': 'doughnut'})
            max_row = len(df_summary) + 1
            chart.add_series({
                'name':       'channel/actual visit',
                'categories': ['Summary', 1, 0, max_row - 1, 0],
                'values':     ['Summary', 1, 3, max_row - 1, 3],
                'data_labels': {'percentage': True},
            })
            chart.set_title({'name': 'channel/actual visit'})
            worksheet_summary.insert_chart('H2', chart)

            # TAB 2: Full SKU Shelf Breakdown for ALL Channels
            sku_cols = [c for c in df.columns if str(c).startswith('架上情況 [')]
            all_sku_details = []
            
            for channel in df_summary['Channel']:
                df_ch = df[df['店鋪_clean'].str.contains(channel, na=False)] if channel != 'SMKT' else df[df['店鋪_clean'].isin(['Wellcome', 'ParkNshop'])]
                tot_v = len(df_ch)
                if tot_v == 0:
                    continue
                    
                for col in sku_cols:
                    clean_sku = col.replace('架上情況 [', '').replace(']', '').strip()
                    col_s = df_ch[col].astype(str)
                    
                    has_stock = col_s.str.contains('有貨').sum()
                    oos_tag = col_s.str.contains('缺貨').sum()
                    no_tag = col_s.str.contains('無貨').sum()
                    cov = round((has_stock / tot_v) * 100, 1)
                    
                    all_sku_details.append({
                        "Channel": channel,
                        "Product SKU": clean_sku,
                        "有貨有牌仔": has_stock,
                        "缺貨有牌仔": oos_tag,
                        "無貨無牌仔 / 無貨唔牌仔": no_tag,
                        "Selling Coverage (%)": cov
                    })
            
            df_all_skus = pd.DataFrame(all_sku_details)
            df_all_skus.to_excel(writer, sheet_name='SKU Breakdown', index=False)

        st.download_button(
            label="🟢 Download Complete Excel Report (.xlsx)",
            data=output.getvalue(),
            file_name=f"Market_Visit_Summary_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")