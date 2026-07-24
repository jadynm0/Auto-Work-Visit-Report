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
        # 1. READ SURVEY DATA FIRST (DYNAMIC SOURCE OF TRUTH)
        # ---------------------------------------------------------
        df_raw = pd.read_excel(uploaded_file, sheet_name='表格回應 1')
        df_raw.columns = df_raw.iloc[0]
        df = df_raw.iloc[1:].reset_index(drop=True)
        df['店鋪_clean'] = df['店鋪'].astype(str).str.strip()
        
        # Get all actual store channels present in the survey
        actual_store_counts = df['店鋪_clean'].value_counts()
        
        # ---------------------------------------------------------
        # 2. READ TARGETS DYNAMICALLY
        # ---------------------------------------------------------
        df_targets = pd.read_excel(uploaded_file, sheet_name='Targets')
        df_targets.columns = [str(c).strip() for c in df_targets.columns]
        
        ch_col = df_targets.columns[0]
        tg_col = df_targets.columns[1]

        df_targets = df_targets.dropna(subset=[ch_col])
        
        targets_dict = {}
        for _, row in df_targets.iterrows():
            ch_raw = str(row[ch_col]).strip()
            ch_name = '7-11' if any(k in ch_raw for k in ['07-11', '7-11', '7/11', '2026-07-11']) else ch_raw
            try:
                targets_dict[ch_name] = int(float(row[tg_col]))
            except:
                targets_dict[ch_name] = 0

        # Combine channels from both Targets sheet AND actual Survey data dynamically
        all_channels = list(dict.fromkeys(list(targets_dict.keys()) + list(actual_store_counts.index)))

        # ---------------------------------------------------------
        # 3. BUILD DYNAMIC SUMMARY TABLE
        # ---------------------------------------------------------
        summary_rows = []
        for channel in all_channels:
            if '7-11' in channel or '07-11' in channel:
                actual = sum(count for store, count in actual_store_counts.items() if any(k in store for k in ['7-11', '7/11', '07-11']))
                channel_label = '7-11'
            else:
                actual = actual_store_counts.get(channel, 0)
                channel_label = channel

            target = targets_dict.get(channel_label, 0)
            
            # Avoid duplicate channel entries
            if any(r['Channel'] == channel_label for r in summary_rows):
                continue

            status = "Target Met 🟢" if actual >= target and target > 0 else "MISSED TARGET ❌"
            
            summary_rows.append({
                "Channel": channel_label,
                "Target Visit": target,
                "Actual Visit": actual,
                "Status": status
            })

        df_summary = pd.DataFrame(summary_rows)

        # ---------------------------------------------------------
        # SECTION 1: OVERALL DASHBOARD & PLOTLY DONUT CHART
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
        # SECTION 2: INTERACTIVE "BY CHANNEL" ANALYSIS
        # ---------------------------------------------------------
        st.divider()
        st.header("🔍 Interactive Analysis By Channel")
        
        channel_options = df_summary['Channel'].tolist()
        selected_ch = st.selectbox("By channel:", options=channel_options)

        df_ch = df[df['店鋪_clean'].str.contains(selected_ch, na=False)]
        
        ch_target = targets_dict.get(selected_ch, 0)
        ch_actual = len(df_ch)

        m1, m2 = st.columns(2)
        m1.metric("Target Visit", ch_target)
        m2.metric("Actual Visit", ch_actual)

        # ---------------------------------------------------------
        # CHOICE COVERAGE RANGE MATRIX
        # ---------------------------------------------------------
        st.subheader(f"📊 {selected_ch} - Choice Coverage Breakdown")
        
        sku_cols = [c for c in df_ch.columns if str(c).startswith('架上情況 [')]
        
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

            # SKU SHELF STATUS TABLE
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
            st.info("No visit records found for this selected channel.")

        # ---------------------------------------------------------
        # EXCEL EXPORT
        # ---------------------------------------------------------
        st.divider()
        st.header("📥 Download Excel Summary")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            
            workbook = writer.book
            worksheet_summary = writer.sheets['Summary']
            
            chart = workbook.add_chart({'type': 'doughnut'})
            max_row = len(df_summary) + 1
            chart.add_series({
                'name':       'channel/actual visit',
                'categories': ['Summary', 1, 0, max_row - 1, 0],
                'values':     ['Summary', 1, 2, max_row - 1, 2],
                'data_labels': {'percentage': True},
            })
            chart.set_title({'name': 'channel/actual visit'})
            worksheet_summary.insert_chart('G2', chart)

            # DEDICATED CHANNEL SHEETS
            for channel in df_summary['Channel']:
                df_ch = df[df['店鋪_clean'].str.contains(channel, na=False)]
                tot_v = len(df_ch)
                if tot_v == 0:
                    continue
                
                sheet_title = channel.replace(':', '').replace('/', '-')[:30]
                
                # Detailed SKU Status Table
                channel_sku_details = []
                for col in sku_cols:
                    clean_sku = col.replace('架上情況 [', '').replace(']', '').strip()
                    col_s = df_ch[col].astype(str)
                    
                    has_stock = col_s.str.contains('有貨').sum()
                    oos_tag = col_s.str.contains('缺貨').sum()
                    no_tag = col_s.str.contains('無貨').sum()
                    cov = round((has_stock / tot_v) * 100, 1)
                    
                    channel_sku_details.append({
                        "Product SKU": clean_sku,
                        "有貨有牌仔": has_stock,
                        "缺貨有牌仔": oos_tag,
                        "無貨無牌仔 / 無貨唔牌仔": no_tag,
                        "Selling Coverage (%)": f"{cov}%"
                    })
                
                df_sku = pd.DataFrame(channel_sku_details)
                df_sku.to_excel(writer, sheet_name=sheet_title, index=False)

        st.download_button(
            label="🟢 Download Multi-Tab Excel Report (.xlsx)",
            data=output.getvalue(),
            file_name=f"Market_Visit_Summary_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")