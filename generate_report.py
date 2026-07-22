import pandas as pd
import glob
import os
import sys

def get_latest_excel_file():
    """Scans the directory for the most recently modified .xlsx file."""
    # Look for files matching '行場計劃*.xlsx' first
    xlsx_files = glob.glob('行場計劃*.xlsx')
    
    # Fallback: look for any .xlsx file (ignoring temporary files starting with ~$)
    if not xlsx_files:
        xlsx_files = [f for f in glob.glob('*.xlsx') if not os.path.basename(f).startswith('~$')]
        
    if not xlsx_files:
        print("❌ No .xlsx files found in the current directory!")
        print("Please ensure your Excel file is saved as .xlsx (not .numbers) in this folder.")
        sys.exit(1)
        
    # Get the file with the latest modification timestamp
    latest_file = max(xlsx_files, key=os.path.getmtime)
    print(f"📁 Auto-detected latest report file: {latest_file}\n")
    return latest_file


def run_report_generator():
    print("Loading data files...\n")
    
    # Dynamically select the newest .xlsx file in the folder
    FILE_PATH = get_latest_excel_file()
    
    # ---------------------------------------------------------
    # STEP 1: DYNAMICALLY LOAD TARGETS FROM "Targets" SHEET
    # ---------------------------------------------------------
    try:
        df_targets = pd.read_excel(FILE_PATH, sheet_name='Targets')
        # Clean up any potential empty rows or trailing spaces
        df_targets = df_targets.dropna(subset=['Channel', 'Target'])
        df_targets['Channel'] = df_targets['Channel'].astype(str).str.strip()
        TARGETS = dict(zip(df_targets['Channel'], df_targets['Target']))
    except Exception as e:
        print(f"❌ Error loading 'Targets' sheet from {FILE_PATH}: {e}")
        print("Please ensure the file is saved as .xlsx (not .numbers) and contains a 'Targets' sheet.")
        sys.exit(1)

    # ---------------------------------------------------------
    # STEP 2: LOAD AND CLEAN RAW SURVEY DATA
    # ---------------------------------------------------------
    try:
        df_raw = pd.read_excel(FILE_PATH, sheet_name='表格回應 1')
        df_raw.columns = df_raw.iloc[0]
        df = df_raw.iloc[1:].reset_index(drop=True)
    except Exception as e:
        print(f"❌ Error loading '表格回應 1' sheet: {e}")
        sys.exit(1)

    # Calculate actual store counts per channel
    store_counts = df['店鋪'].value_counts()

    # Aggregate channels to match target KPIs
    actual_smkt = store_counts.get('Wellcome', 0) + store_counts.get('ParkNshop', 0)
    actual_min_chain = store_counts.get('Aeon', 0) + store_counts.get("city'super", 0)

    actuals = {
        '7-11': store_counts.get('7-11', 0),
        'Circle K': store_counts.get('Circle K', 0),
        'SMKT': actual_smkt,
        'Min.Chain': actual_min_chain,
        '佳寶': store_counts.get('佳寶', 0)
    }

    # ---------------------------------------------------------
    # STEP 3: GENERATE SUMMARY REPORT
    # ---------------------------------------------------------
    districts = ", ".join(df['地區'].dropna().unique().tolist())

    report = f"""==================================================
MONTHLY MARKET VISIT PERFORMANCE REPORT
==================================================
[Overview Summary]
- Total Regions Checked : {df['地區'].dropna().nunique()} districts
- Covered Regions       : {districts}
- Total Stores Audited  : {len(df)} KA Branches

[KPI Performance Tracking (Target vs Actual)]
"""

    for channel, target in TARGETS.items():
        actual = actuals.get(channel, 0)
        status = "Target Met 🟢" if actual >= target else "MISSED TARGET ❌"
        report += f"- {channel:<10}: Goal {target:>3} stores | Actual: {actual:>3} stores -> ({status})\n"

    # Helper function to compute product availability metrics
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

    # Add specific store group summaries
    report += analyze_channel_products(df[df['店鋪'] == '7-11'], "7-11", "Freshly Made")
    report += analyze_channel_products(df[df['店鋪'] == 'Circle K'], "Circle K", "Fresh & Pre-packaged")

    print(report)

if __name__ == '__main__':
    run_report_generator()