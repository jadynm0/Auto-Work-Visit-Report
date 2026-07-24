import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io

st.set_page_config(page_title="Market Visit Performance Report", page_icon="📊", layout="wide")
st.title("📊 Monthly Market Visit Performance Report Generator")

uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    st.success("File uploaded successfully! Processing report...")
    
    try:
        # ---------------------------------------------------------
        # 1. READ TARGETS & NORMALIZE CHANNEL NAMES
        # ---------------------------------------------------------
        df_targets = pd.read_excel(uploaded_file, sheet_name='Targets')
        df_targets.columns = [str(c).strip() for c in df_targets.columns]
        
        ch_col = df_targets.columns[0]
        tg_col = df_targets.columns[1]

        df_targets = df_targets.dropna(subset=[ch_col, tg_col])
        
        cleaned_targets = {}
        for _, row in df_targets.iterrows():
            ch_raw = str(row[ch_col]).strip()
            
            if any(k in ch_raw for k in ['07-11', '7-11', '7/11', '2026-07-11']):
                ch_name = '7-11'
            else:
                ch_name = ch_raw
                
            try:
                target_val = int(float(row[tg_col]))
            except:
                target_val = 0
                
            cleaned_targets[ch_name] = target_val

        # ---------------------------------------------------------
        # 2. READ SURVEY DATA & COUNT STORES
        # ---------------------------------------------------------
        df_raw = pd.read_excel(uploaded_file, sheet_name='表格回應 1')
        df_raw.columns = df_raw.iloc[0]
        df = df_raw.iloc[1:].reset_index(drop=True)
        
        df['店鋪_clean'] = df['店鋪'].astype(str).str.strip()
        store_counts = df['店鋪_clean'].value_counts()

        summary_data = []
        for channel, target in cleaned_targets.items():
            if channel == '7-11':
                actual = sum(count for store, count in store_counts.items() if any(k in store for k in ['7-11', '7/11', '07-11']))
            elif channel == 'SMKT':
                actual = store_counts.get('Wellcome', 0) + store_counts.get('ParkNshop', 0)
            elif channel.lower() in ['min.chain', 'min chain']:
                actual = store_counts.get('Aeon', 0) + store_counts.get("city'super", 0)
            else:
                actual = store_counts.get(channel, 0)

            status = "Target Met 🟢" if actual >= target else "MISSED TARGET ❌"
            summary_data.append({
                "Channel": channel,
                "Target Visit": target,
                "Actual Visit": actual,
                "Status": status
            })

        df_summary = pd.DataFrame(summary_data)

        # ---------------------------------------------------------
        # 3. DISPLAY WEB DASHBOARD WITH EXACT PIE/DONUT CHART
        # ---------------------------------------------------------
        st.divider()
        st.header("📌 Channel Visit Summary & Share")
        
        col_left, col_right = st.columns([1, 1])
        
        with col_left:
            st.subheader("Visit Performance Summary")
            st.dataframe(df_summary, use_container_width=True)
            st.metric("Total Stores Audited", len(df))

        with col_right:
            st.subheader("channel/actual visit Share")
            
            # Filter for non-zero visits
            chart_df = df_summary[df_summary['Actual Visit'] > 0]
            
            if not chart_df.empty:
                # Render Pie Chart matching Gg's sheet layout
                fig, ax = plt.subplots(figsize=(6, 6))
                
                # Make figure background transparent to blend with Streamlit theme
                fig.patch.set_alpha(0.0)
                
                wedges, texts, autotexts = ax.pie(
                    chart_df['Actual Visit'], 
                    labels=chart_df['Channel'], 
                    autopct='%1.1f%%',
                    startangle=140,
                    textprops=dict(color="white"),
                    wedgeprops=dict(width=0.4, edgecolor='none') # Donut shape
                )
                
                plt.setp(autotexts, size=10, weight="bold")
                ax.axis('equal')
                
                st.pyplot(fig)

        # ---------------------------------------------------------
        # 4. GENERATE DOWNLOADABLE EXCEL WORKBOOK
        # ---------------------------------------------------------
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Summary']
            
            chart = workbook.add_chart({'type': 'doughnut'})
            max_row = len(df_summary) + 1
            chart.add_series({
                'name':       'channel/actual visit',
                'categories': ['Summary', 1, 0, max_row - 1, 0],
                'values':     ['Summary', 1, 2, max_row - 1, 2],
                'data_labels': {'percentage': True},
            })
            
            chart.set_title({'name': 'channel/actual visit'})
            chart.set_style(10)
            worksheet.insert_chart('F2', chart)

        excel_data = output.getvalue()

        # ---------------------------------------------------------
        # 5. RENDER DOWNLOAD BUTTONS & TEXT SUMMARY
        # ---------------------------------------------------------
        st.divider()
        st.header("📥 Download Options")
        
        st.download_button(
            label="🟢 Download Excel Summary (.xlsx)",
            data=excel_data,
            file_name=f"Market_Visit_Summary_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        districts = ", ".join(df['地區'].dropna().astype(str).str.strip().unique().tolist())
        report = f"""==================================================
MONTHLY MARKET VISIT PERFORMANCE REPORT
==================================================
Processed File : {uploaded_file.name}

[Overview Summary]
- Total Regions Checked : {df['地區'].dropna().nunique()} districts
- Covered Regions       : {districts}
- Total Stores Audited  : {len(df)} KA Branches

[KPI Performance Tracking (Target vs Actual)]
"""
        for _, row in df_summary.iterrows():
            report += f"- {row['Channel']:<10}: Goal {row['Target Visit']:>3} stores | Actual: {row['Actual Visit']:>3} stores -> ({row['Status']})\n"

        st.subheader("📋 Executive Text Summary")
        st.text_area("Copyable Report Text", value=report, height=300)

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")