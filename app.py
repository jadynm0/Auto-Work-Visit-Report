import streamlit as st
import pandas as pd
import plotly.express as px
import re
import io

st.set_page_config(page_title="Market Visit Summary Dashboard", page_icon="📊", layout="wide")
st.title("📊 Monthly Market Visit Performance Dashboard")

uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    st.success("File uploaded successfully! Processing summary...")
    
    try:
        # ---------------------------------------------------------
        # 1. READ RAW SURVEY DATA DYNAMICALLY
        # ---------------------------------------------------------
        df_raw = pd.read_excel(uploaded_file, sheet_name='表格回應 1')
        df_raw.columns = [str(c).strip() for c in df_raw.iloc[0]]
        df = df_raw.iloc[1:].reset_index(drop=True)
        
        df['店鋪_clean'] = df['店鋪'].astype(str).str.strip()
        
        # Extract ALL product columns dynamically
        sku_cols = [c for c in df.columns if '架上情況 [' in str(c)]
        
        def clean_sku_name(col_name):
            match = re.search(r'\[(.*?)\]', str(col_name))
            return match.group(1).strip() if match else str(col_name).strip()

        sku_mapping = {col: clean_sku_name(col) for col in sku_cols}

        # ---------------------------------------------------------
        # 2. READ TARGETS DYNAMICALLY
        # ---------------------------------------------------------
        targets_dict = {}
        try:
            df_targets = pd.read_excel(uploaded_file, sheet_name='Targets')
            df_targets.columns = [str(c).strip() for c in df_targets.columns]
            
            ch_col, tg_col = df_targets.columns[0], df_targets.columns[1]
            df_targets = df_targets.dropna(subset=[ch_col])
            
            for _, row in df_targets.iterrows():
                ch_raw = str(row[ch_col]).strip()
                ch_name = '7-11' if any(k in ch_raw for k in ['07-11', '7-11', '7/11', '2026-07-11']) else ch_raw
                try:
                    targets_dict[ch_name] = int(float(row[tg_col]))
                except:
                    targets_dict[ch_name] = 0
        except Exception:
            st.warning("Targets sheet not found or unreadable; proceeding with dynamic visit counts.")

        # ---------------------------------------------------------
        # 3. BUILD DYNAMIC CHANNEL SUMMARY
        # ---------------------------------------------------------
        raw_channels = df['店鋪_clean'].unique().tolist()
        
        consolidated_counts = {}
        for ch in raw_channels:
            if any(k in str(ch) for k in ['7-11', '7/11', '07-11']):
                label = '7-11'
            else:
                label = ch
            
            count = (df['店鋪_clean'] == ch).sum()
            consolidated_counts[label] = consolidated_counts.get(label, 0) + count

        summary_rows = []
        for ch_name, actual_count in consolidated_counts.items():
            tg_val = targets_dict.get(ch_name, 0)
            status = "Target Met 🟢" if actual_count >= tg_val and tg_val > 0 else ("No Target Set ⚪" if tg_val == 0 else "MISSED TARGET ❌")
            
            summary_rows.append({
                "Channel": ch_name,
                "Target Visit": tg_val,
                "Actual Visit": actual_count,
                "Status": status
            })

        df_summary = pd.DataFrame(summary_rows)

        # ---------------------------------------------------------
        # SECTION 1: OVERALL KPI DASHBOARD & PLOTLY DONUT
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
            fig = px.pie(chart_df, values='Actual Visit', names='Channel', hole=0.5, color_discrete_sequence=px.colors.qualitative.Set2)
            fig.update_traces(textinfo='percent+label', textposition='outside')
            st.plotly_chart(fig, use_container_width=True)

        # ---------------------------------------------------------
        # SECTION 2: INTERACTIVE "BY CHANNEL" OR "ALL STORES (OVERALL)"
        # ---------------------------------------------------------
        st.divider()
        st.header("🔍 Interactive Analysis (Channel or Overall)")
        
        channel_options = ["All Stores (Overall)"] + df_summary['Channel'].tolist()
        selected_ch = st.selectbox("Select View:", options=channel_options)

        if selected_ch == "All Stores (Overall)":
            df_ch = df
            st.subheader("🌐 Overall (All Brands / Shops Combined)")
            st.metric("Total Stores Audited", len(df_ch))
        else:
            df_ch = df[df['店鋪_clean'].str.contains(selected_ch, na=False)]
            ch_target = targets_dict.get(selected_ch, 0)
            ch_actual = len(df_ch)

            m1, m2 = st.columns(2)
            m1.metric("Target Visit", ch_target)
            m2.metric("Actual Visit", ch_actual)

        # CHOICE COVERAGE RANGE MATRIX (0, 1-4, 5-9, >9)
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

            # DETAILED SKU SHELF STATUS TABLE
            st.subheader(f"🛒 {selected_ch} - Detailed SKU Shelf Status")
            sku_details = []
            for orig_col, clean_name in sku_mapping.items():
                col_s = df_ch[orig_col].astype(str)
                
                has_stock = col_s.str.contains('有貨').sum()
                oos_tag = col_s.str.contains('缺貨').sum()
                no_tag = col_s.str.contains('無貨').sum()
                cov = round((has_stock / tot_v) * 100, 1)
                
                sku_details.append({
                    "Product SKU": clean_name,
                    "有貨有牌仔": has_stock,
                    "缺貨有牌仔": oos_tag,
                    "無貨無牌仔 / 無貨唔牌仔": no_tag,
                    "Product Coverage (%)": f"{cov}%"
                })
            st.dataframe(pd.DataFrame(sku_details), use_container_width=True)
        else:
            st.info("No visit records found.")

        # ---------------------------------------------------------
        # SECTION 3: MULTI-TAB EXCEL EXPORT WITH CHOICE COVERAGE & AUTO-WIDTHS
        # ---------------------------------------------------------
        st.divider()
        st.header("📥 Download Complete Formatted Excel Summary")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Custom Excel Formats
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'border': 1, 'align': 'center'})
            cell_fmt = workbook.add_format({'border': 1, 'align': 'center'})
            
            # 1. Summary Sheet
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            ws_summary = writer.sheets['Summary']
            
            for col_idx, col in enumerate(df_summary.columns):
                max_len = max(df_summary[col].astype(str).map(len).max(), len(str(col))) + 5
                ws_summary.set_column(col_idx, col_idx, max(max_len, 12), cell_fmt)
                ws_summary.write(0, col_idx, col, header_fmt)

            chart = workbook.add_chart({'type': 'doughnut'})
            max_row = len(df_summary) + 1
            chart.add_series({
                'name':       'channel/actual visit',
                'categories': ['Summary', 1, 0, max_row - 1, 0],
                'values':     ['Summary', 1, 2, max_row - 1, 2],
                'data_labels': {'percentage': True},
            })
            chart.set_title({'name': 'channel/actual visit'})
            ws_summary.insert_chart('G2', chart)

            # Helper function to write Choice Coverage Matrix + Detailed SKU Table to any Excel sheet
            def export_detailed_channel_sheet(sub_df, sheet_name):
                tot_visits = len(sub_df)
                if tot_visits == 0:
                    return

                # Calculate Choice Coverage Range Matrix
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

                # Build SKU Detailed Table
                sku_records = []
                for orig_col, clean_name in sku_mapping.items():
                    col_series = sub_df[orig_col].astype(str)
                    has_stock = col_series.str.contains('有貨').sum()
                    oos_tag = col_series.str.contains('缺貨').sum()
                    no_tag = col_series.str.contains('無貨').sum()
                    cov = round((has_stock / tot_visits) * 100, 1) if tot_visits > 0 else 0
                    
                    sku_records.append({
                        "Product SKU": clean_name,
                        "有貨有牌仔": has_stock,
                        "缺貨有牌仔": oos_tag,
                        "無貨無牌仔 / 無貨唔牌仔": no_tag,
                        "Coverage (%)": f"{cov}%"
                    })
                df_sku_details = pd.DataFrame(sku_records)

                # Write Matrix at row 0
                df_choice_matrix.to_excel(writer, sheet_name=sheet_name, startrow=0, index=False)
                
                # Write SKU details table starting at row 7
                df_sku_details.to_excel(writer, sheet_name=sheet_name, startrow=7, index=False)

                ws = writer.sheets[sheet_name]

                # Format Choice Matrix headers
                for col_idx, col in enumerate(df_choice_matrix.columns):
                    ws.write(0, col_idx, col, header_fmt)

                # Format SKU Details headers (row 7)
                for col_idx, col in enumerate(df_sku_details.columns):
                    ws.write(7, col_idx, col, header_fmt)

                # Auto-adjust column widths
                for col_idx, col in enumerate(df_sku_details.columns):
                    max_len = max(
                        df_sku_details[col].astype(str).map(len).max(),
                        len(str(col))
                    ) + 5
                    ws.set_column(col_idx, col_idx, max(max_len, 15), cell_fmt)

            # 2. Overall Sheet (All Stores Combined)
            export_detailed_channel_sheet(df, 'All Stores (Overall)')

            # 3. Dynamic Sheets per Channel
            for ch_label in df_summary['Channel']:
                sub_df = df[df['店鋪_clean'].str.contains(ch_label, na=False)]
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