import streamlit as st
import pandas as pd

# Page Configuration
st.set_page_config(page_title="Market Visit Performance Report", page_icon="📊", layout="wide")

st.title("📊 Monthly Market Visit Performance Report Generator")

uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    st.success("File uploaded successfully! Processing report...")
    
    try:
        # ---------------------------------------------------------
        # 1. READ TARGETS & FIX DATE CONVERSION BUG (7-11 -> 2026-07-11)
        # ---------------------------------------------------------
        df_targets = pd.read_excel(uploaded_file, sheet_name='Targets')
        df_targets.columns = [str(c).strip() for c in df_targets.columns]
        
        ch_col = df_targets.columns[0]
        tg_col = df_targets.columns[1]

        df_targets = df_targets.dropna(subset=[ch_col, tg_col])
        
        # CLEANING: Fix Excel auto-converting '7-11' into datetime objects
        cleaned_targets = {}
        for _, row in df_targets.iterrows():
            ch_raw = str(row[ch_col]).strip()
            # If pandas read '7-11' as a date (e.g. '2026-07-11 00:00:00' or '07-11')
            if '07-11' in ch_raw or '7-11' in ch_raw or '7/11' in ch_raw:
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
        
        # Count actual store visits clean
        store_counts = df['店鋪'].astype(str).str.strip().value_counts()

        # ---------------------------------------------------------
        # 3. BUILD EXECUTIVE REPORT TEXT MATCHING GG'S SHEET
        # ---------------------------------------------------------
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

        for channel, target in cleaned_targets.items():
            actual = store_counts.get(channel, 0)
            status = "Target Met 🟢" if actual >= target else "MISSED TARGET ❌"
            report += f"- {channel:<10}: Goal {target:>3} stores | Actual: {actual:>3} stores -> ({status})\n"

        # Function to generate SKU tables for specific channels
        def analyze_channel_products(df_store, store_name):
            if len(df_store) == 0:
                return f"\n[{store_name} - Product Coverage Summary]\n  * No visit data recorded.\n"
                
            report_section = f"\n[{store_name} - Product Coverage Summary (Total Audited: {len(df_store)} stores)]\n"
            cols = [c for c in df_store.columns if str(c).startswith('架上情況 [')]
            
            for col in cols:
                clean_sku = col.replace('架上情況 [', '').replace(']', '').strip()
                total_stores = len(df_store)
                
                col_str = df_store[col].astype(str)
                in_stock = col_str.str.contains('有貨').sum()
                oos = col_str.str.contains('缺貨').sum()
                rate = (in_stock / total_stores) * 100 if total_stores > 0 else 0
                
                if in_stock > 0 or oos > 0:
                    report_section += f"  * {clean_sku:<18} | On-Shelf: {in_stock:>2} | OOS: {oos:>2} | Coverage: {rate:>5.1f}%\n"
            return report_section

        # Generate breakdowns for 7-11 & Circle K
        report += analyze_channel_products(df[df['店鋪'].astype(str).str.contains('7-11')], "7-11")
        report += analyze_channel_products(df[df['店鋪'].astype(str).str.contains('Circle K')], "Circle K")

        # ---------------------------------------------------------
        # 4. RENDER IN STREAMLIT UI
        # ---------------------------------------------------------
        st.subheader("📋 Corrected Executive Report Summary")
        st.text_area("Copyable Report Text", value=report, height=500)
        
        st.download_button(
            label="💾 Download Report (.txt)",
            data=report,
            file_name=f"Market_Visit_Report_{uploaded_file.name.replace('.xlsx', '')}.txt",
            mime="text/plain"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")