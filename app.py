import streamlit as st
import pandas as pd

# Page Configuration
st.set_page_config(page_title="Market Visit Report Generator", page_icon="📊", layout="wide")

st.title("📊 Monthly Market Visit Performance Report Generator")
st.write("Upload your monthly market visit Excel file to generate a structured report instantly.")

# File Uploader Widget
uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    st.success("File uploaded successfully! Processing report...")
    
    try:
        # ---------------------------------------------------------
        # STEP 1: READ TARGETS SHEET (TYPE-SAFE)
        # ---------------------------------------------------------
        df_targets = pd.read_excel(uploaded_file, sheet_name='Targets')
        
        # Clean column headers
        df_targets.columns = [str(c).strip() for c in df_targets.columns]
        col_map = {str(c).lower(): c for c in df_targets.columns}
        
        if 'channel' in col_map and 'target' in col_map:
            channel_col = col_map['channel']
            target_col = col_map['target']
        else:
            channel_col = df_targets.columns[0]
            target_col = df_targets.columns[1]

        df_targets = df_targets.dropna(subset=[channel_col, target_col])
        df_targets[channel_col] = df_targets[channel_col].astype(str).str.strip()
        
        # Ensure Target values are strictly numeric integers
        df_targets[target_col] = pd.to_numeric(df_targets[target_col], errors='coerce').fillna(0).astype(int)
        
        TARGETS = dict(zip(df_targets[channel_col], df_targets[target_col]))
        
        # ---------------------------------------------------------
        # STEP 2: READ SURVEY DATA SHEET
        # ---------------------------------------------------------
        df_raw = pd.read_excel(uploaded_file, sheet_name='表格回應 1')
        df_raw.columns = df_raw.iloc[0]
        df = df_raw.iloc[1:].reset_index(drop=True)
        
        # Calculate actual counts
        store_counts = df['店鋪'].value_counts()
        actual_smkt = int(store_counts.get('Wellcome', 0) + store_counts.get('ParkNshop', 0))
        actual_min_chain = int(store_counts.get('Aeon', 0) + store_counts.get("city'super", 0))

        actuals = {
            '7-11': int(store_counts.get('7-11', 0)),
            'Circle K': int(store_counts.get('Circle K', 0)),
            'SMKT': actual_smkt,
            'Min.Chain': actual_min_chain,
            '佳寶': int(store_counts.get('佳寶', 0))
        }

        # ---------------------------------------------------------
        # STEP 3: BUILD REPORT TEXT
        # ---------------------------------------------------------
        districts = ", ".join(df['地區'].dropna().unique().tolist())

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

        for channel, raw_target in TARGETS.items():
            actual = int(actuals.get(channel, 0))
            target = int(raw_target)
            status = "Target Met 🟢" if actual >= target else "MISSED TARGET ❌"
            report += f"- {channel:<10}: Goal {target:>3} stores | Actual: {actual:>3} stores -> ({status})\n"

        def analyze_channel_products(df_store, store_name, item_type="Fresh"):
            report_section = f"\n[{store_name} - {item_type} Product Coverage Summary]\n"
            cols = [c for c in df_store.columns if str(c).startswith('架上情況 [')]
            
            for col in cols:
                clean_name = col.replace('架上情況 [', '').replace(']', '')
                total_stores = len(df_store)
                if total_stores == 0: 
                    continue
                    
                in_stock = df_store[col].astype(str).str.contains('有貨').sum()
                oos = df_store[col].astype(str).str.contains('缺貨').sum()
                rate = (in_stock / total_stores) * 100
                
                if in_stock > 0 or oos > 0:
                    report_section += f"  * {clean_name:<15} | On-Shelf: {in_stock:>2} | OOS: {oos:>2} | Coverage: {rate:>5.1f}%\n"
            return report_section

        report += analyze_channel_products(df[df['店鋪'] == '7-11'], "7-11", "Freshly Made")
        report += analyze_channel_products(df[df['店鋪'] == 'Circle K'], "Circle K", "Fresh & Pre-packaged")

        # ---------------------------------------------------------
        # STEP 4: DISPLAY IN WEB UI
        # ---------------------------------------------------------
        st.subheader("📋 Generated Report Output")
        st.text_area("Copyable Report Text", value=report, height=450)
        
        st.download_button(
            label="💾 Download Report (.txt)",
            data=report,
            file_name=f"Market_Visit_Report_{uploaded_file.name.replace('.xlsx', '')}.txt",
            mime="text/plain"
        )
        
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
        st.info("Please make sure your file is a valid .xlsx file containing both '表格回應 1' and 'Targets' sheets.")