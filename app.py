import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import xlsxwriter
import re
import io
from datetime import datetime

st.set_page_config(page_title="Market Visit Summary Dashboard", page_icon="📊", layout="wide")
st.title("📊 Monthly Market Visit Performance & AI Strategic Dashboard")

# Initialize persistent in-app session storage
if 'stored_surveys' not in st.session_state:
    st.session_state['stored_surveys'] = {}
if 'stored_pl_deals' not in st.session_state:
    st.session_state['stored_pl_deals'] = {}

# Dual File Uploaders
col_up1, col_up2 = st.columns(2)
with col_up1:
    uploaded_files = st.file_uploader(
        "1. Upload Market Visit Survey File(s) (.xlsx) [Single or Multiple Months]", 
        type=["xlsx"], 
        accept_multiple_files=True
    )
with col_up2:
    uploaded_pl = st.file_uploader(
        "2. Upload Product Listing File (.xlsx) [Optional]", 
        type=["xlsx"]
    )

# 1. Standard Fallback KA Monthly Targets
DEFAULT_KA_TARGETS = {
    '7-11': 100,
    'Circle K': 40,
    'Wellcome': 20,
    'ParkNshop': 20,
    '佳寶': 20,
    'Aeon': 10,
    "city'super": 10,
    'UNY': 0
}

# 2. 18 HK Administrative Districts Grouped by Region
HK_REGION_GROUPS = {
    'Hong Kong Island (港島)': ['中西區', '東區', '南區', '灣仔區'],
    'Kowloon (九龍)': ['九龍城區', '觀塘區', '深水埗區', '黃大仙區', '油尖旺區'],
    'New Territories & Islands (新界及離島)': ['離島區', '葵青區', '北區', '西貢區', '沙田區', '大埔區', '荃灣區', '屯門區', '元朗區']
}

# 3. Standard Client Name Map for Bilingual Alignment
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

# 4. Product SKU Name Cross-Reference
SKU_NAME_MAP = {
    '310ml楊枝甘露(常溫)': ['常溫楊枝甘露310ml', '芒果椰果柚子甘露310ml'],
    '蘆楊': ['蘆薈楊枝甘露', '蘆薈楊枝甘露450ml'],
    '紫米露': ['椰香紫米甘露', '椰香紫米甘露 330g'],
    '爆檸蜜': ['凍檸蜜'],
    '花膠冰糖雪耳甘露飲品': ['花膠雪耳冰糖甘露', '花膠雪耳冰糖甘露330g'],
    '杏仁露': ['杏仁露 330ml'],
    '無添加糖豆乳': ['無添加糖豆乳飲品500ml', '無添加糖豆漿']
}

# 5. Product Category Classification
CHILLED_KEYWORDS = ['蘆楊', '杏仁露', '紫米露', '奶茶', '竹笙', '雪梨川貝', '火麻仁', '紅豆沙', '綠豆沙', '火麻仁拿鐵', '黑豆黑芝麻', '花膠', '無添加糖豆乳']
OTHER_KEYWORDS = ['湯', '豬腳薑', '龜苓膏']

def get_product_category(sku_name):
    s = str(sku_name).strip()
    if any(k in s for k in OTHER_KEYWORDS):
        return 'Others (Soups & Food / 湯品及其他)'
    elif any(k in s for k in CHILLED_KEYWORDS):
        return 'Chilled Beverages & Desserts (鮮製飲品及甜品)'
    else:
        return 'Room-Temperature Beverages (常溫/預製飲品)'

def format_reporting_month(m_key):
    """Converts 202606 -> ('June 2026', '2026年6月')"""
    m_match = re.search(r'(\d{4})(\d{2})', str(m_key))
    if m_match:
        year, month = m_match.group(1), m_match.group(2)
        dt = datetime.strptime(f"{year}{month}", "%Y%m")
        return dt.strftime("%B %Y"), f"{year}年{int(month)}月"
    return str(m_key), str(m_key)

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
            if has_p:
                return True
        if clean_s == clean_p or clean_s in clean_p or clean_p in clean_s:
            if has_p:
                return True
    return False

def parse_targets_dynamically(df_targets):
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

def generate_dynamic_market_insights(df, df_summary, df_dist_summary, sku_mapping, m_label_zh):
    total_audits = len(df)
    visited_dist = len(df_dist_summary[df_dist_summary['Status'] == 'Visited 🟢'])
    
    met_channels = []
    missed_channels = []
    for _, r in df_summary[df_summary['Channel'] != 'Total'].iterrows():
        ch = r['Channel']
        act = int(r['Actual Visit']) if str(r['Actual Visit']).isdigit() else 0
        tg = int(r['Target Visit']) if str(r['Target Visit']).isdigit() else 0
        if tg > 0:
            pct = round((act / tg) * 100, 1)
            if act >= tg:
                met_channels.append(f"{ch}（{pct}% / 實際{act}間）")
            else:
                missed_channels.append(f"{ch}（{pct}% / 目標{tg}間/實際{act}間）")
        elif act > 0:
            met_channels.append(f"{ch}（走訪{act}間）")

    sku_stats = []
    for orig_col, clean_name in sku_mapping.items():
        col_s = df[orig_col].astype(str)
        in_stk = col_s.str.contains('有貨').sum()
        oos = col_s.str.contains('缺貨').sum()
        cov = round((in_stk / total_audits) * 100, 1) if total_audits > 0 else 0
        cat = get_product_category(clean_name)
        sku_stats.append({
            'name': clean_name, 'in_stock': in_stk, 'oos': oos, 'cov': cov, 'category': cat
        })
        
    df_sku_stat = pd.DataFrame(sku_stats)
    top_stars = df_sku_stat.sort_values(by='cov', ascending=False).head(5)
    problem_skus = df_sku_stat[df_sku_stat['cov'] < 10.0].sort_values(by='cov', ascending=True)
    oos_alerts = df_sku_stat[df_sku_stat['oos'] > 0].sort_values(by='oos', ascending=False).head(5)

    ch_profiles = {}
    for ch in ['7-11', 'Circle K', 'Wellcome', 'ParkNshop', '佳寶', 'Aeon', "city'super"]:
        sub = df[df['店鋪_clean'].str.contains(ch, regex=False, na=False)]
        if len(sub) > 0:
            ch_skus = []
            for orig_col, clean_name in sku_mapping.items():
                s_cnt = sub[orig_col].astype(str).str.contains('有貨').sum()
                rate = round((s_cnt / len(sub)) * 100, 1)
                ch_skus.append((clean_name, s_cnt, rate))
            ch_skus.sort(key=lambda x: x[2], reverse=True)
            ch_profiles[ch] = {
                'count': len(sub),
                'top_items': ch_skus[:3],
                'bottom_items': [x for x in ch_skus if x[2] < 15.0][:3],
                'zero_count': len([x for x in ch_skus if x[2] == 0])
            }

    text = f"""### 📊 一、 全局市場分析與洞察 ({m_label_zh} Market Analysis & Findings)

#### 1. 通路巡查執行力與目標達成 (Audit Execution & Objectives)
* **整體走訪覆蓋**: {m_label_zh}實地走訪各大零售通路共 **{total_audits} 間分店**，全面覆蓋全港 **{visited_dist}/18 個行政區**。
* **重點通路執行情況**:
  * **超額/順利完成通路**: {('、'.join(met_channels)) if met_channels else '各通路穩定走訪'}。
  * **存在缺口/進度落後通路**: {('、'.join(missed_channels)) if missed_channels else '無明顯落後通路，整體外勤執行度良好'}。

#### 2. 各主要通路現況表現 (Channel Listing Performance)
"""
    if '7-11' in ch_profiles or 'Circle K' in ch_profiles:
        text += "* **便利店 (7-Eleven / Circle K)**:\n"
        if '7-11' in ch_profiles:
            p = ch_profiles['7-11']
            text += f"  * **7-11** ({p['count']} 間分店): 主力熱銷為 **{p['top_items'][0][0]}** ({p['top_items'][0][2]}%) 及 **{p['top_items'][1][0]}** ({p['top_items'][1][2]}%)；而部分品項表現疲弱（如 {', '.join([x[0] for x in p['bottom_items']])}）。\n"
        if 'Circle K' in ch_profiles:
            p = ch_profiles['Circle K']
            text += f"  * **Circle K** ({p['count']} 間分店): 表現突出為 **{p['top_items'][0][0]}** ({p['top_items'][0][2]}%) 及 **{p['top_items'][1][0]}** ({p['top_items'][1][2]}%)；另有 {p['zero_count']} 款產品未見陳列 (0%)。\n"

    if 'Wellcome' in ch_profiles or 'ParkNshop' in ch_profiles:
        text += "* **大眾超級市場 (Wellcome 惠康 / ParkNshop 百佳)**:\n"
        if 'Wellcome' in ch_profiles:
            p = ch_profiles['Wellcome']
            text += f"  * **惠康 (Wellcome)** ({p['count']} 間分店): 常溫與 1L 家庭裝表現亮眼，以 **{p['top_items'][0][0]}** ({p['top_items'][0][2]}%) 與 **{p['top_items'][1][0]}** ({p['top_items'][1][2]}%) 居首。\n"
        if 'ParkNshop' in ch_profiles:
            p = ch_profiles['ParkNshop']
            text += f"  * **百佳 (PNS)** ({p['count']} 間分店): 預製涼茶覆蓋率極高，皇牌為 **{p['top_items'][0][0]}** ({p['top_items'][0][2]}%) 及 **{p['top_items'][1][0]}** ({p['top_items'][1][2]}%)。\n"

    if '佳寶' in ch_profiles or 'Aeon' in ch_profiles or "city'super" in ch_profiles:
        text += "* **平價賣場與精品百貨 (佳寶 / Aeon / city'super)**:\n"
        if '佳寶' in ch_profiles:
            p = ch_profiles['佳寶']
            text += f"  * **佳寶**: 核心常溫涼茶款鋪貨率高達 90%-100%，以極致性價比產品為主。\n"
        if 'Aeon' in ch_profiles:
            p = ch_profiles['Aeon']
            text += f"  * **Aeon**: 鮮製系列（甜品及特色飲品）表現最為平均且覆蓋率優異。\n"
        if "city'super" in ch_profiles:
            p = ch_profiles["city'super"]
            text += f"  * **city'super**: 高端客群偏好預製樽裝，短保鮮期鮮製產品專區仍有提升空間。\n"

    text += f"""
#### 3. 明星商品、告急品項與痛點分析 (Product Performance & Problem Child SKUs)
* **明星產品 (Star & Cash Cow SKUs)**:
  * 全港覆蓋率頂尖品項：{', '.join([f"**{r['name']}** ({r['cov']}%)" for _, r in top_stars.iterrows()])}。
  * 夏季「蘆薈楊枝甘露」及 1L 大容量家庭裝（五花茶/甘蔗汁）持續扮演帶動整體業績的核心火車頭。
* **告急產品與痛點 (Problem Child SKUs)**:
  * 部分低轉速或特定通路未上架品項（如 {', '.join([f"{r['name']} ({r['cov']}%)" for _, r in problem_skus.head(4).iterrows()])}）上架率偏低，面臨邊緣化風險。
  * 超市湯品系列及部分長尾糖水受季節性影響，陳列面較為薄弱。
* **缺貨有牌仔 (OOS) 警示**:
  * {', '.join([f"**{r['name']}** (共 {r['oos']} 間分店標籤缺貨)" for _, r in oos_alerts.iterrows()]) if len(oos_alerts) > 0 else '全線通路供貨穩定，未見異常缺貨情況'}。

---

### 🎯 二、 通路策略與行銷建議 ({m_label_zh} Strategic Recommendations)

#### 1. SKU 結構調整與資源重組 (SKU Optimization)
* **【引爆明星】擴大「楊枝甘露」與 1L 家庭裝效應**:
  * 將核心熱銷品作為主打，於便利店推行「鮮製甜品 + 即飲涼茶」加價購優惠；在超市則強化 1L 家庭裝（五花茶、甘蔗汁）多件促銷以拉高客單價。
* **【止蝕與拯救】告急產品專案處置**:
  * 針對便利店上架率極低的品項，檢討產品定位與保質期；若非供應鏈斷貨，建議考慮縮減陳列面或替換為高轉速品項（如將位置讓給紫米露或綠豆沙）。
* **【季節性品項提前佈局】**:
  * 針對超市陳列較弱的湯品系列，建議提前於 8-9 月規劃「初秋滋補/潤燥暖湯」主題陳列專區，搭配套裝優惠提振銷量。

#### 2. 通路差異化行銷策略 (Channel-Specific Strategy)
* **便利店 (7-11 / Circle K)**:
  * 聚焦「即飲、解渴、甜品補給」訴求，確保冷櫃第一層高週轉面，並優先防範熱賣糖水缺貨（OOS）。
* **大眾超市 (Wellcome / PNS)**:
  * 聚焦「家庭囤貨、健康養生」訴求，主推 1L 家庭裝與 4 入裝預製涼茶，並爭取將高人氣鮮製糖水擴展至更多分店冷櫃。
* **平價賣場 (佳寶) & 精品超市 (Aeon / city'super)**:
  * 佳寶維持 4 款核心高性價比涼茶穩定供貨；Aeon 可作為高單價及新口味甜品（如黑豆、花膠系列）試水溫基地。

#### 3. 業務執行與陳列優化 (Sales & Trade Marketing Execution)
* **解決「缺貨 (OOS)」與「無陳列 (Absent)」門市問題**:
  * 業務團隊應根據本期巡查數據與各區店長核對訂單，特別針對出現缺貨牌仔的熱賣款要求即時補貨，確保冰櫃架面數（Shelf-Space）。
* **動態平衡 18 區巡查路線**:
  * 持續追蹤未達標通路及未走訪行政區，適度輪替跨區巡查資源，確保每月審查數據全面無死角。
"""
    return text

# Pre-calculate store counts for channel matrix to prevent f-string backslash errors
def get_channel_counts(df):
    c_store = len(df[df['店鋪_clean'].str.contains('7-11|Circle K|OK', regex=True, na=False)])
    smkt = len(df[df['店鋪_clean'].str.contains('Wellcome|惠康|ParkNshop|百佳', regex=True, na=False)])
    kb = len(df[df['店鋪_clean'].str.contains('佳寶', regex=False, na=False)])
    prem = len(df[df['店鋪_clean'].str.contains("Aeon|city'super", regex=True, na=False)])
    return c_store, smkt, kb, prem

# ---------------------------------------------------------
# PROCESS PRODUCT LISTING
# ---------------------------------------------------------
if uploaded_pl is not None:
    try:
        df_pl = pd.read_excel(uploaded_pl, sheet_name=0)
        clients_raw = [str(c).replace('\n', ' ').strip() for c in df_pl.iloc[3, 2:].values if pd.notnull(c)]
        parsed_deals = {}
        for r_idx in range(4, len(df_pl)):
            sku_name = str(df_pl.iloc[r_idx, 1]).strip()
            if sku_name and sku_name != 'nan' and 'sub total' not in sku_name.lower():
                for c_idx, raw_client in enumerate(clients_raw):
                    if (c_idx + 2) < len(df_pl.columns):
                        val = str(df_pl.iloc[r_idx, c_idx + 2]).strip().upper()
                        std_channel = raw_client
                        for std_ch, aliases in CLIENT_NAME_MAP.items():
                            if any(a.lower() in raw_client.lower() or raw_client.lower() in a.lower() for a in aliases):
                                std_channel = std_ch
                                break
                        if std_channel not in parsed_deals:
                            parsed_deals[std_channel] = {}
                        
                        if val == 'P':
                            parsed_deals[std_channel][sku_name] = True
                        elif sku_name not in parsed_deals[std_channel]:
                            parsed_deals[std_channel][sku_name] = False
        st.session_state['stored_pl_deals'] = parsed_deals
        st.toast("Product listing deals loaded & stored in memory!", icon="📜")
    except Exception as e:
        st.error(f"Error parsing product listing: {e}")

# ---------------------------------------------------------
# PROCESS AND PERSIST SURVEY WORKBOOKS IN IN-APP STORAGE
# ---------------------------------------------------------
if uploaded_files:
    for u_file in uploaded_files:
        try:
            excel_obj = pd.ExcelFile(u_file)
            s_name = next((s for s in excel_obj.sheet_names if any(k in str(pd.read_excel(u_file, sheet_name=s, nrows=2).values) for k in ['店鋪', '姓名', '時間', '地區'])), excel_obj.sheet_names[0])
            raw_d = pd.read_excel(u_file, sheet_name=s_name)
            
            if any(k in " ".join([str(c) for c in raw_d.iloc[0].values]) for k in ['店鋪', '姓名', '地區']):
                raw_d.columns = [str(c).strip() for c in raw_d.iloc[0]]
                df_m = raw_d.iloc[1:].reset_index(drop=True)
            else:
                df_m = raw_d.reset_index(drop=True)
                
            m_match = re.search(r'2026\d{2}', u_file.name)
            m_label = m_match.group(0) if m_match else u_file.name[:10]
            st.session_state['stored_surveys'][m_label] = (df_m, u_file.name, u_file.getvalue())
        except Exception:
            pass

# Manage in-app storage via an expander
if st.session_state['stored_surveys']:
    with st.expander(f"🗄️ In-App Historical Storage: {len(st.session_state['stored_surveys'])} Month(s) Loaded", expanded=False):
        stored_tags = [f"`{format_reporting_month(m)[0]}`" for m in sorted(st.session_state['stored_surveys'].keys())]
        st.write("Loaded Month(s): " + ", ".join(stored_tags))
        if st.button("🗑️ Clear All Stored Historical Months"):
            st.session_state['stored_surveys'] = {}
            st.rerun()

# ---------------------------------------------------------
# ACTIVE REPORTING RENDER
# ---------------------------------------------------------
if st.session_state['stored_surveys']:
    sorted_months = sorted(st.session_state['stored_surveys'].keys())
    
    month_options_map = {m: f"{format_reporting_month(m)[0]} ({format_reporting_month(m)[1]} - {m})" for m in sorted_months}
    selected_display = st.selectbox(
        "📅 Select Active Reporting Month / Year:", 
        options=list(month_options_map.values()), 
        index=len(sorted_months)-1
    )
    
    active_month = next(k for k, v in month_options_map.items() if v == selected_display)
    active_m_en, active_m_zh = format_reporting_month(active_month)
    
    df, active_fname, active_bytes = st.session_state['stored_surveys'][active_month]
    product_listing_deals = st.session_state.get('stored_pl_deals', {})

    # Data cleaning
    df['店鋪_clean'] = df['店鋪'].astype(str).str.strip()
    df['姓名_clean'] = df['姓名'].astype(str).str.strip()
    df['地區_clean'] = df['地區'].apply(normalize_district_18)
    
    sku_cols = [c for c in df.columns if '架上情況 [' in str(c)]
    def clean_sku_name(col_name):
        match = re.search(r'\[(.*?)\]', str(col_name))
        return match.group(1).strip() if match else str(col_name).strip()
    sku_mapping = {col: clean_sku_name(col) for col in sku_cols}

    # Dynamic target parsing with DEFAULT_KA_TARGETS fallback
    targets_dict = DEFAULT_KA_TARGETS.copy()
    total_shops_dict = {}
    try:
        excel_obj = pd.ExcelFile(io.BytesIO(active_bytes))
        target_sheets = [s for s in excel_obj.sheet_names if any(k in s.lower() for k in ['target', 'summary', '目標'])]
        if target_sheets:
            df_t_raw = pd.read_excel(io.BytesIO(active_bytes), sheet_name=target_sheets[0])
            parsed_tg, parsed_hk = parse_targets_dynamically(df_t_raw)
            if parsed_tg:
                targets_dict.update(parsed_tg)
            if parsed_hk:
                total_shops_dict.update(parsed_hk)
    except Exception:
        pass

    # District coverage
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
    df_dist_summary = pd.DataFrame(district_summary_rows).fillna('N/A')

    # Channel summary
    salespeople = [str(s).strip() for s in df['姓名_clean'].unique() if pd.notnull(s) and str(s).strip().lower() not in ['nan', 'none', '']]
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
    df_summary = pd.DataFrame(summary_rows).fillna('N/A')

    ai_generated_text = generate_dynamic_market_insights(df, df_summary, df_dist_summary, sku_mapping, active_m_zh)
    c_cnt, s_cnt, k_cnt, p_cnt = get_channel_counts(df)

    # ---------------------------------------------------------
    # MAIN DASHBOARD TABS
    # ---------------------------------------------------------
    tab_dash, tab_trend, tab_ai = st.tabs(["📊 Performance Dashboard", "📈 Multi-Month Trend Analysis", "💡 AI Strategic Recommendations"])

    with tab_dash:
        st.header(f"📌 {active_m_en} ({active_m_zh}) Market Visit Summary")
        m_col1, m_col2 = st.columns([1, 2])
        m_col1.metric("Total Stores Audited", len(df))
        m_col2.metric("HK 18-District Coverage", f"{visited_count_18}/18 Administrative Districts Audited")

        with st.expander("📍 View Regional & District Breakdown (港島 / 九龍 / 新界及離島)", expanded=False):
            st.dataframe(df_dist_summary, use_container_width=True)

        col1, col2 = st.columns([1.4, 1])
        with col1:
            st.subheader("Channel & Salesperson Breakdown")
            st.dataframe(df_summary, use_container_width=True)
        with col2:
            st.subheader("Channel Audit Share")
            chart_df = df_summary[(df_summary['Channel'] != 'Total') & (pd.to_numeric(df_summary['Actual Visit'], errors='coerce') > 0)]
            if len(chart_df) > 0:
                fig = px.pie(chart_df, values='Actual Visit', names='Channel', hole=0.5, color_discrete_sequence=px.colors.qualitative.Set2)
                fig.update_traces(textinfo='percent+label', textposition='outside')
                st.plotly_chart(fig, use_container_width=True)

        st.divider()
        st.header("🔍 Interactive Channel Analysis")
        valid_channels = [c for c in df_summary['Channel'].tolist() if c != 'Total' and int(df_summary.loc[df_summary['Channel']==c, 'Actual Visit'].values[0]) > 0]
        channel_options = ["All Stores (Overall)"] + valid_channels
        selected_ch = st.selectbox("Select View:", options=channel_options)

        df_ch = df if selected_ch == "All Stores (Overall)" else df[df['店鋪_clean'].str.contains(selected_ch, regex=False, na=False)]
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
        ]).fillna('N/A')
        st.dataframe(df_choice, use_container_width=True)

        # SKU Availability Table
        sku_details = []
        client_deals = product_listing_deals.get(selected_ch, {})
        for orig_col, clean_name in sku_mapping.items():
            col_s = df_ch[orig_col].astype(str)
            has_stock = col_s.str.contains('有貨').sum()
            oos_tag = col_s.str.contains('缺貨').sum()
            no_tag = col_s.str.contains('無貨').sum()
            cov = round((has_stock / tot_v) * 100, 1) if tot_v > 0 else 0
            cat = get_product_category(clean_name)
            
            if selected_ch == "All Stores (Overall)":
                deal_note = "Aggregated Overall"
                should_include = True
            else:
                has_deal = match_sku_deal(clean_name, client_deals)
                if has_stock > 0 and has_deal is True:
                    deal_note = "With Deal"
                    should_include = True
                elif has_stock > 0 and has_deal is False:
                    deal_note = "Unlisted / Stocked without Deal (No Deal but in stock)"
                    should_include = True
                elif has_stock == 0 and has_deal is True:
                    deal_note = "With Deal (OOS)"
                    should_include = True
                else:
                    deal_note = "0% due to No Deal (Not Listed)"
                    should_include = False

            if should_include:
                sku_details.append({
                    "Category": cat,
                    "Product SKU": clean_name,
                    "Deal Status": deal_note,
                    "有貨有牌仔": has_stock,
                    "缺貨有牌仔": oos_tag,
                    "無貨無牌仔": no_tag,
                    "Coverage (%)": f"{cov}%",
                    "Performance Rating": get_coverage_rating(cov)
                })
        df_sku_view = pd.DataFrame(sku_details).fillna('N/A')
        st.dataframe(df_sku_view, use_container_width=True)

        # Category Stock Charts
        if selected_ch != "All Stores (Overall)" and len(df_sku_view) > 0:
            st.subheader(f"📈 {selected_ch} - Category Stock Charts")
            cat_tabs = st.tabs(["🍵 Room-Temperature (常溫)", "🍧 Chilled & Desserts (鮮製及甜品)", "🍲 Others & Soups (湯品及其他)"])
            cat_map_names = {
                "🍵 Room-Temperature (常溫)": 'Room-Temperature Beverages (常溫/預製飲品)',
                "🍧 Chilled & Desserts (鮮製及甜品)": 'Chilled Beverages & Desserts (鮮製飲品及甜品)',
                "🍲 Others & Soups (湯品及其他)": 'Others (Soups & Food / 湯品及其他)'
            }
            for tab_obj, (tab_title, cat_name) in zip(cat_tabs, cat_map_names.items()):
                with tab_obj:
                    sub_cat_df = df_sku_view[df_sku_view['Category'] == cat_name]
                    if len(sub_cat_df) > 0:
                        fig_bar = go.Figure(data=[
                            go.Bar(name='有貨有牌仔 (In-Stock)', x=sub_cat_df['Product SKU'], y=sub_cat_df['有貨有牌仔'], marker_color='#2ca02c'),
                            go.Bar(name='缺貨有牌仔 (OOS)', x=sub_cat_df['Product SKU'], y=sub_cat_df['缺貨有牌仔'], marker_color='#d62728'),
                            go.Bar(name='無貨無牌仔 (No Stock)', x=sub_cat_df['Product SKU'], y=sub_cat_df['無貨無牌仔'], marker_color='#7f7f7f')
                        ])
                        fig_bar.update_layout(barmode='stack', xaxis_tickangle=-45, height=420)
                        st.plotly_chart(fig_bar, use_container_width=True)

    with tab_trend:
        st.header("📈 Multi-Month Historical Trend Tracking (2026 MoM Analytics)")
        if len(sorted_months) > 1:
            trend_data = []
            for m_key in sorted_months:
                d_m, _, _ = st.session_state['stored_surveys'][m_key]
                t_aud = len(d_m)
                m_label_display = format_reporting_month(m_key)[0]
                row_trend = {"Month": m_label_display, "Total Audits": t_aud}
                for hero in ['蘆楊', '夏枯草', '雞骨草', '竹蔗茅根', '咸柑桔', '紫米露', '紅豆沙', '綠豆沙', '五花茶1L']:
                    matching_c = [c for c in d_m.columns if hero in c]
                    if matching_c:
                        in_cnt = d_m[matching_c[0]].astype(str).str.contains('有貨').sum()
                        row_trend[hero] = round((in_cnt / t_aud) * 100, 1) if t_aud > 0 else 0
                trend_data.append(row_trend)
                
            df_trend = pd.DataFrame(trend_data)
            st.subheader("1. Monthly Audit Volume Progression (2026)")
            fig_vol = px.bar(df_trend, x='Month', y='Total Audits', text='Total Audits', color='Month', color_discrete_sequence=px.colors.qualitative.Pastel)
            st.plotly_chart(fig_vol, use_container_width=True)
            
            st.subheader("2. Core Hero SKU Coverage % Month-over-Month Evolution")
            hero_skus = [c for c in df_trend.columns if c not in ['Month', 'Total Audits']]
            fig_line = px.line(df_trend, x='Month', y=hero_skus, markers=True)
            fig_line.update_layout(yaxis_title="Coverage Rate (%)", height=500)
            st.plotly_chart(fig_line, use_container_width=True)
            
            st.dataframe(df_trend, use_container_width=True)
        else:
            st.info("💡 Upload another month's survey file (e.g. upload June 2026, then upload July 2026) to generate automated multi-month trend charts.")

    # ---------------------------------------------------------
    # TAB: REDESIGNED EXECUTIVE AI VIEWS & RECOMMENDATIONS
    # ---------------------------------------------------------
    with tab_ai:
        st.header(f"💡 AI Executive Strategic Summary ({active_m_en} / {active_m_zh})")
        
        sku_calc_list = []
        for orig_col, clean_name in sku_mapping.items():
            col_s = df[orig_col].astype(str)
            in_s = col_s.str.contains('有貨').sum()
            oos_s = col_s.str.contains('缺貨').sum()
            cov_r = round((in_s / len(df)) * 100, 1) if len(df) > 0 else 0
            sku_calc_list.append({'name': clean_name, 'in_stock': in_s, 'oos': oos_s, 'cov': cov_r})
        df_calc = pd.DataFrame(sku_calc_list)
        top_stars_web = df_calc.sort_values(by='cov', ascending=False).head(5)
        top_oos_web = df_calc[df_calc['oos'] > 0].sort_values(by='oos', ascending=False).head(5)

        # 1. Top KPI Metric Cards
        kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
        with kpi_col1:
            st.metric("Total Audited Stores", f"{len(df)} 間分店")
        with kpi_col2:
            st.metric("18-District Reach", f"{visited_count_18}/18 行政區")
        with kpi_col3:
            star_name = top_stars_web.iloc[0]['name'] if len(top_stars_web) > 0 else "N/A"
            star_cov = top_stars_web.iloc[0]['cov'] if len(top_stars_web) > 0 else 0
            st.metric("Top Hero Product", f"{star_name} ({star_cov}%)")
        with kpi_col4:
            oos_name = top_oos_web.iloc[0]['name'] if len(top_oos_web) > 0 else "None"
            oos_cnt = top_oos_web.iloc[0]['oos'] if len(top_oos_web) > 0 else 0
            st.metric("Top OOS Risk Item", f"{oos_name} ({oos_cnt} stores)")

        st.divider()

        # 2. Side-by-Side Visual Charts
        st.subheader("📊 Executive Visual Insights & Channel Performance")
        ai_c1, ai_c2 = st.columns(2)
        with ai_c1:
            ch_chart_data = df_summary[df_summary['Channel'] != 'Total'].copy()
            ch_chart_data['Target Visit'] = pd.to_numeric(ch_chart_data['Target Visit'], errors='coerce').fillna(0)
            ch_chart_data['Actual Visit'] = pd.to_numeric(ch_chart_data['Actual Visit'], errors='coerce').fillna(0)
            
            fig_ai_ch = go.Figure(data=[
                go.Bar(name='目標走訪 (Target)', x=ch_chart_data['Channel'], y=ch_chart_data['Target Visit'], marker_color='#B8CCE4'),
                go.Bar(name='實際走訪 (Actual)', x=ch_chart_data['Channel'], y=ch_chart_data['Actual Visit'], marker_color='#366092')
            ])
            fig_ai_ch.update_layout(barmode='group', title="各大通路目標 vs. 實際走訪數量", height=380)
            st.plotly_chart(fig_ai_ch, use_container_width=True)

        with ai_c2:
            fig_ai_star = px.bar(
                top_stars_web.sort_values(by='cov', ascending=True), 
                x='cov', y='name', orientation='h', text='cov',
                title="⭐ 皇牌品項 TOP 5 全港覆蓋率 (%)",
                color_discrete_sequence=['#4F81BD']
            )
            fig_ai_star.update_traces(texttemplate='%{text}%', textposition='outside')
            fig_ai_star.update_layout(xaxis_title="覆蓋率 (%)", yaxis_title="", height=380)
            st.plotly_chart(fig_ai_star, use_container_width=True)

        st.divider()

        # 3. Channel Strategy Matrix Table
        st.subheader("🎯 各通路差異化現況與營運策略矩陣 (Channel Strategy Matrix)")

        matrix_rows = [
            {
                "通路類型": "便利店 (7-Eleven / Circle K)",
                "走訪分店數": f"{c_cnt} 間",
                "主力熱銷品項": "鮮製甜品 (蘆楊、紫米露、紅綠豆沙) 及 500ml 夏枯草",
                "弱勢/缺貨品項": "豆漿低覆蓋 (4%)；火麻仁拿鐵滯銷；紅綠豆沙及蘆楊局部缺貨",
                "核心營運與進貨建議": "主攻「即飲消暑與甜品補給」；推行加價購優惠；及時補足甜品冷櫃架面防止缺貨"
            },
            {
                "通路類型": "大眾超市 (Wellcome / PNS)",
                "走訪分店數": f"{s_cnt} 間",
                "主力熱銷品項": "預製涼茶 (夏枯草、雞骨草、竹蔗茅根) 及 1L 家庭裝 (五花茶/甘蔗汁)",
                "弱勢/缺貨品項": "湯品系列 (20%-38%) 及長尾鮮製糖水滲透率偏低",
                "核心營運與進貨建議": "主打「家庭囤貨與日常養生」；強化 1L 家庭裝多件促銷；提前於 8-9 月規劃初秋暖湯滋補陳列"
            },
            {
                "通路類型": "平價賣場 (佳寶)",
                "走訪分店數": f"{k_cnt} 間",
                "主力熱銷品項": "核心 4 款預製涼茶 (夏枯草、雞骨草、竹蔗、咸柑桔達 90%+)",
                "弱勢/缺貨品項": "銀菊露 (約 8%-24%) 鋪貨疲弱",
                "核心營運與進貨建議": "鎖定「極致性價比」；集中資源確保 4 大熱銷款安全庫存，避免長尾滯銷款佔用資金"
            },
            {
                "通路類型": "精品/日系百貨 (Aeon / city'super)",
                "走訪分店數": f"{p_cnt} 間",
                "主力熱銷品項": "Aeon 鮮製甜品及特色奶茶 (覆蓋率達 85%-100%)",
                "弱勢/缺貨品項": "city'super 鮮製短保專區較小，走訪進度尚有缺口",
                "核心營運與進貨建議": "作為高單價與特色新品 (港式奶茶、黑豆、花膠系列) 試水溫基地；推動擴大冷藏陳列架面"
            }
        ]
        st.dataframe(pd.DataFrame(matrix_rows), use_container_width=True)

        st.divider()

        # 4. Detailed Narrative Section
        st.markdown(ai_generated_text)

    # ---------------------------------------------------------
    # MULTI-TAB EXCEL EXPORT WORKBOOK (WITH CLEAN EXECUTIVE AI FORMAT)
    # ---------------------------------------------------------
    st.divider()
    st.header(f"📥 Download Complete Formatted Excel Report ({active_m_en})")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        workbook = writer.book
        
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'border': 1, 'align': 'center'})
        cell_fmt = workbook.add_format({'border': 1, 'align': 'center'})
        
        ai_banner_fmt = workbook.add_format({
            'bold': True, 'font_size': 14, 'font_color': 'white', 
            'bg_color': '#1F497D', 'align': 'left', 'valign': 'vcenter', 'indent': 1
        })
        ai_section_hdr = workbook.add_format({
            'bold': True, 'font_size': 11, 'font_color': '#1F497D', 
            'bg_color': '#DCE6F1', 'border': 1, 'valign': 'vcenter', 'indent': 1
        })
        ai_subhdr_fmt = workbook.add_format({
            'bold': True, 'font_size': 10, 'font_color': '#1F497D', 'valign': 'vcenter', 'indent': 1
        })
        ai_text_fmt = workbook.add_format({
            'font_size': 9, 'font_color': '#333333', 'text_wrap': True, 'valign': 'top', 'indent': 2
        })
        table_hdr_fmt = workbook.add_format({
            'bold': True, 'font_size': 9, 'font_color': 'white', 'bg_color': '#366092', 'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        table_oos_hdr = workbook.add_format({
            'bold': True, 'font_size': 9, 'font_color': 'white', 'bg_color': '#C00000', 'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        table_cell_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 9})
        table_pct_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 9, 'num_format': '0.0%'})
        table_oos_val = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 9, 'font_color': '#C00000', 'bold': True})
        td_wrap = workbook.add_format({'border': 1, 'font_size': 9, 'valign': 'vcenter', 'text_wrap': True, 'indent': 1})

        # 1. Summary Sheet
        df_summary.to_excel(writer, sheet_name='Summary', index=False)
        ws_summary = writer.sheets['Summary']
        for col_idx, col in enumerate(df_summary.columns):
            max_len = max(df_summary[col].astype(str).map(len).max(), len(str(col))) + 5
            ws_summary.set_column(col_idx, col_idx, max(max_len, 12), cell_fmt)
            ws_summary.write(0, col_idx, col, header_fmt)

        # 2. District Coverage Sheet
        df_dist_summary.to_excel(writer, sheet_name='District Coverage', index=False)
        ws_dist = writer.sheets['District Coverage']
        for col_idx, col in enumerate(df_dist_summary.columns):
            max_len = max(df_dist_summary[col].astype(str).map(len).max(), len(str(col))) + 5
            ws_dist.set_column(col_idx, col_idx, max(max_len, 15), cell_fmt)
            ws_dist.write(0, col_idx, col, header_fmt)

        # ---------------------------------------------------------
        # 3. REDESIGNED EXECUTIVE AI VIEWS & RECOMMENDATIONS SHEET
        # ---------------------------------------------------------
        ws_ai = workbook.add_worksheet('AI Views & Recommendations')
        ws_ai.hide_gridlines(0)
        ws_ai.set_column('A:A', 2)
        ws_ai.set_column('B:B', 20)
        ws_ai.set_column('C:C', 11)
        ws_ai.set_column('D:D', 11)
        ws_ai.set_column('E:E', 12)
        ws_ai.set_column('F:F', 11)
        ws_ai.set_column('G:G', 3)
        ws_ai.set_column('H:H', 22)
        ws_ai.set_column('I:I', 12)
        ws_ai.set_column('J:J', 12)
        ws_ai.set_column('K:K', 14)
        ws_ai.set_column('L:L', 14)
        ws_ai.set_column('M:M', 14)

        # Title Header Banner
        ws_ai.merge_range('B2:M2', f'💡 {active_m_en} ({active_m_zh}) 市場巡查 AI 總結分析與營運策略報告 (Executive Strategic Summary)', ai_banner_fmt)
        ws_ai.set_row(1, 28)

        # Dynamic Section Offsets
        ch_export_df = df_summary[df_summary['Channel'] != 'Total'].copy()
        start_sec1 = 4
        num_ch = len(ch_export_df)
        end_sec1 = start_sec1 + num_ch + 1

        # Section 1: Channel Targets Table (Cols B:F)
        ws_ai.merge_range(f'B{start_sec1}:M{start_sec1}', '📊 一、 通路巡查目標執行達成表與對比圖 (Audit Target vs. Actual)', ai_section_hdr)
        ws_ai.set_row(start_sec1 - 1, 22)

        headers_ch = ['通路名稱 (Channel)', '目標走訪', '實際走訪', '達成率 (%)', '達成狀況']
        for c_i, h in enumerate(headers_ch, start=1):
            ws_ai.write(start_sec1, c_i, h, table_hdr_fmt)
            
        for r_i, (_, row) in enumerate(ch_export_df.iterrows(), start=start_sec1 + 1):
            ch_name = str(row['Channel'])
            tg = int(row['Target Visit']) if str(row['Target Visit']).isdigit() else 0
            act = int(row['Actual Visit']) if str(row['Actual Visit']).isdigit() else 0
            rate_val = (act / tg) if tg > 0 else 0
            st_text = "已達標 🟢" if (tg > 0 and act >= tg) else ("未達標 🔴" if tg > 0 else "無目標 ⚪")
            
            ws_ai.write(r_i, 1, ch_name, table_cell_fmt)
            ws_ai.write(r_i, 2, tg, table_cell_fmt)
            ws_ai.write(r_i, 3, act, table_cell_fmt)
            ws_ai.write(r_i, 4, rate_val, table_pct_fmt)
            ws_ai.write(r_i, 5, st_text, table_cell_fmt)

        # Chart 1: Channel Target vs. Actual on the right (Cols H:M)
        chart_tg_act = workbook.add_chart({'type': 'column'})
        max_ch_row = start_sec1 + num_ch
        chart_tg_act.add_series({
            'name':       ['AI Views & Recommendations', start_sec1, 2],
            'categories': ['AI Views & Recommendations', start_sec1 + 1, 1, max_ch_row, 1],
            'values':     ['AI Views & Recommendations', start_sec1 + 1, 2, max_ch_row, 2],
            'fill':       {'color': '#B8CCE4'}
        })
        chart_tg_act.add_series({
            'name':       ['AI Views & Recommendations', start_sec1, 3],
            'categories': ['AI Views & Recommendations', start_sec1 + 1, 1, max_ch_row, 1],
            'values':     ['AI Views & Recommendations', start_sec1 + 1, 3, max_ch_row, 3],
            'fill':       {'color': '#366092'},
            'data_labels': {'value': True}
        })
        chart_tg_act.set_title({'name': '各大通路目標 vs. 實際走訪分店數'})
        chart_tg_act.set_size({'width': 500, 'height': 210})
        ws_ai.insert_chart(f'H{start_sec1 + 1}', chart_tg_act)

        # Section 2: Top 5 & OOS (Dynamic Starting Row to Prevent Overwrite)
        start_sec2 = max(end_sec1 + 2, 16)
        ws_ai.merge_range(f'B{start_sec2}:M{start_sec2}', '⭐ 二、 皇牌品項 TOP 5 與門市缺貨 (OOS) 預警 (Hero SKUs & Stockout Alert)', ai_section_hdr)
        ws_ai.set_row(start_sec2 - 1, 22)

        # Top 5 Table (Cols B:E)
        ws_ai.merge_range(f'B{start_sec2 + 1}:E{start_sec2 + 1}', '⭐ 皇牌熱賣品項 TOP 5 (全港覆蓋率)', table_hdr_fmt)
        for ci, h in enumerate(['產品 SKU', '走訪有貨分店', '全港覆蓋率', '評級'], start=1):
            ws_ai.write(start_sec2 + 1, ci, h, table_hdr_fmt)
            
        sku_calc = []
        for orig_col, clean_name in sku_mapping.items():
            col_s = df[orig_col].astype(str)
            in_stk = col_s.str.contains('有貨').sum()
            oos_c = col_s.str.contains('缺貨').sum()
            cov_r = (in_stk / len(df)) if len(df) > 0 else 0
            sku_calc.append({'name': clean_name, 'in_stock': in_stk, 'oos': oos_c, 'cov': cov_r})
        df_ai_calc = pd.DataFrame(sku_calc)
        top5_stars = df_ai_calc.sort_values(by='cov', ascending=False).head(5)
        top5_oos = df_ai_calc[df_ai_calc['oos'] > 0].sort_values(by='oos', ascending=False).head(5)

        for idx_k, (_, r_k) in enumerate(top5_stars.iterrows(), start=start_sec2 + 2):
            ws_ai.write(idx_k, 1, r_k['name'], table_cell_fmt)
            ws_ai.write(idx_k, 2, int(r_k['in_stock']), table_cell_fmt)
            ws_ai.write(idx_k, 3, float(r_k['cov']), table_pct_fmt)
            ws_ai.write(idx_k, 4, get_coverage_rating(r_k['cov'] * 100), table_cell_fmt)

        # Chart 2: Top Star SKUs Coverage Bar Chart (Cols H:M)
        chart_top5 = workbook.add_chart({'type': 'bar'})
        chart_top5.add_series({
            'name':       '全港覆蓋率',
            'categories': ['AI Views & Recommendations', start_sec2 + 2, 1, start_sec2 + 6, 1],
            'values':     ['AI Views & Recommendations', start_sec2 + 2, 3, start_sec2 + 6, 3],
            'fill':       {'color': '#4F81BD'},
            'data_labels': {'value': True, 'num_format': '0.0%'}
        })
        chart_top5.set_title({'name': 'Top 5 皇牌品項全港覆蓋率 (%)'})
        chart_top5.set_legend({'none': True})
        chart_top5.set_size({'width': 500, 'height': 170})
        ws_ai.insert_chart(f'H{start_sec2 + 1}', chart_top5)

        # OOS Table below Top 5
        start_oos_row = start_sec2 + 8
        ws_ai.merge_range(f'B{start_oos_row}:E{start_oos_row}', '⚠️ 門市缺貨有牌仔 (OOS) 預警名單', table_oos_hdr)
        for ci, h in enumerate(['產品 SKU', '缺貨店鋪數', '補貨急迫度', '跟進行動'], start=1):
            ws_ai.write(start_oos_row, ci, h, table_oos_hdr)

        for idx_o, (_, r_o) in enumerate(top5_oos.iterrows(), start=start_oos_row + 1):
            ws_ai.write(idx_o, 1, r_o['name'], table_cell_fmt)
            ws_ai.write(idx_o, 2, int(r_o['oos']), table_oos_val)
            urgency = '緊急 🔴' if r_o['oos'] >= 8 else ('高 🟠' if r_o['oos'] >= 5 else '中 🟡')
            action = '即時安排車隊補貨' if r_o['oos'] >= 8 else ('核對店長採購配額' if r_o['oos'] >= 5 else '常態排班巡查補位')
            ws_ai.write(idx_o, 3, urgency, table_cell_fmt)
            ws_ai.write(idx_o, 4, action, table_cell_fmt)

        # Section 3: Channel Strategy Matrix
        start_sec3 = start_oos_row + len(top5_oos) + 2
        ws_ai.merge_range(f'B{start_sec3}:M{start_sec3}', '🎯 三、 各主要通路現況與營運策略矩陣 (Channel Strategy Matrix)', ai_section_hdr)
        ws_ai.set_row(start_sec3 - 1, 22)

        ws_ai.write(start_sec3, 1, '通路類別', table_hdr_fmt)
        ws_ai.write(start_sec3, 2, '走訪分店', table_hdr_fmt)
        ws_ai.merge_range(start_sec3, 3, start_sec3, 4, '主力熱銷品項', table_hdr_fmt)
        ws_ai.merge_range(start_sec3, 5, start_sec3, 6, '弱勢/缺貨品項', table_hdr_fmt)
        ws_ai.merge_range(start_sec3, 7, start_sec3, 12, '核心營運與進貨建議', table_hdr_fmt)

        matrix_content = [
            ('便利店 (7-Eleven / Circle K)', f"{c_cnt} 間",
             '鮮製甜品 (蘆楊、紫米露、紅綠豆沙) 及 500ml 夏枯草', '豆漿低覆蓋 (4%)；火麻仁拿鐵滯銷；紅綠豆沙及蘆楊局部缺貨',
             '主攻「即飲消暑與甜品補給」；推行加價購優惠；及時補足甜品冷櫃架面防止缺貨'),
            ('大眾超市 (Wellcome / PNS)', f"{s_cnt} 間",
             '預製涼茶 (夏枯草、雞骨草、竹蔗茅根) 及 1L 家庭裝', '超市湯品系列 (20%-38%) 及長尾鮮製糖水滲透率偏低',
             '主打「家庭囤貨與日常養生」；強化 1L 家庭裝多件促銷；提前於 8-9 月規劃初秋暖湯滋補陳列'),
            ('平價賣場 (佳寶)', f"{k_cnt} 間",
             '核心 4 款預製涼茶 (夏枯草、雞骨草、竹蔗、咸柑桔達 90%+)', '銀菊露 (約 8%-24%) 鋪貨疲弱',
             '鎖定「極致性價比」；集中資源確保 4 大熱銷款安全庫存，避免長尾滯銷款佔用資金'),
            ('精品百貨 (Aeon / city\'super)', f"{p_cnt} 間",
             'Aeon 鮮製甜品及特色奶茶 (覆蓋率達 85%-100%)', 'city\'super 鮮製短保專區較小，走訪進度尚有缺口',
             '作為高單價與特色新品 (港式奶茶、黑豆、花膠系列) 試水溫基地；推動擴大冷藏陳列架面')
        ]

        curr_r = start_sec3 + 1
        for row_c in matrix_content:
            ws_ai.write(curr_r, 1, row_c[0], table_cell_fmt)
            ws_ai.write(curr_r, 2, row_c[1], table_cell_fmt)
            ws_ai.merge_range(curr_r, 3, curr_r, 4, row_c[2], td_wrap)
            ws_ai.merge_range(curr_r, 5, curr_r, 6, row_c[3], td_wrap)
            ws_ai.merge_range(curr_r, 7, curr_r, 12, row_c[4], td_wrap)
            ws_ai.set_row(curr_r, 28)
            curr_r += 1

        # Section 4: Narrative Insights
        curr_r += 2
        for raw_line in ai_generated_text.split('\n'):
            line = raw_line.strip()
            if not line or line.startswith('---'):
                continue
            
            clean_str = line.replace('**', '').replace('###', '').replace('####', '').strip()
            if clean_str.startswith('#'):
                clean_str = clean_str.lstrip('# ').strip()

            if line.startswith('### 📊 一、') or line.startswith('### 🎯 二、'):
                curr_r += 1
                ws_ai.merge_range(curr_r, 1, curr_r, 12, clean_str, ai_section_hdr)
                ws_ai.set_row(curr_r, 22)
                curr_r += 1
            elif line.startswith('#### '):
                ws_ai.write(curr_r, 1, clean_str, ai_subhdr_fmt)
                curr_r += 1
            elif line.startswith('* **') or line.startswith('* '):
                bullet_str = '• ' + clean_str.lstrip('*•- ').strip()
                ws_ai.merge_range(curr_r, 1, curr_r, 12, bullet_str, ai_text_fmt)
                ws_ai.set_row(curr_r, 18)
                curr_r += 1
            else:
                ws_ai.merge_range(curr_r, 1, curr_r, 12, clean_str, ai_text_fmt)
                curr_r += 1

        # 4. Multi-Month Historical Trend Sheet (Automatic when >= 2 months loaded)
        if len(sorted_months) > 1:
            ws_mom = workbook.add_worksheet('MoM Trend Analysis')
            ws_mom.write('A1', '📈 1. Monthly Store Audit Progression by Channel (MoM)', workbook.add_format({'bold': True, 'font_size': 13, 'font_color': '#1F497D'}))
            
            ch_list = ['7-11', 'Circle K', 'Wellcome', 'ParkNshop', '佳寶', 'Aeon', "city'super", 'UNY', 'Total']
            ch_mom_records = []
            for ch in ch_list:
                rec = {'Channel': ch}
                for m_key in sorted_months:
                    d_m, _, _ = st.session_state['stored_surveys'][m_key]
                    m_name = format_reporting_month(m_key)[0]
                    if ch == 'Total':
                        cnt = len(d_m)
                    else:
                        cnt = len(d_m[d_m['店鋪'].astype(str).str.contains(ch, regex=False, na=False)])
                    rec[f'{m_name} Audits'] = cnt
                ch_mom_records.append(rec)
            df_ch_mom_exp = pd.DataFrame(ch_mom_records)
            df_ch_mom_exp.to_excel(writer, sheet_name='MoM Trend Analysis', startrow=2, index=False)

            start_sku_row = len(df_ch_mom_exp) + 5
            ws_mom.write(start_sku_row - 1, 0, '🛒 2. Core Hero SKU Shelf Coverage % Evolution (MoM)', workbook.add_format({'bold': True, 'font_size': 13, 'font_color': '#1F497D'}))
            
            hero_skus_list = ['蘆楊', '夏枯草', '雞骨草', '竹蔗茅根', '咸柑桔', '紫米露', '紅豆沙', '綠豆沙', '五花茶1L', '甘蔗汁1L', '杏仁露']
            sku_mom_records = []
            for hero in hero_skus_list:
                s_rec = {'Product SKU': hero}
                for m_key in sorted_months:
                    d_m, _, _ = st.session_state['stored_surveys'][m_key]
                    m_name = format_reporting_month(m_key)[0]
                    matching_c = [c for c in d_m.columns if hero in c]
                    cov_pct = 0.0
                    if matching_c and len(d_m) > 0:
                        in_c = d_m[matching_c[0]].astype(str).str.contains('有貨').sum()
                        cov_pct = round((in_c / len(d_m)) * 100, 1)
                    s_rec[f'{m_name} Coverage'] = f"{cov_pct}%"
                sku_mom_records.append(s_rec)
            df_sku_mom_exp = pd.DataFrame(sku_mom_records)
            df_sku_mom_exp.to_excel(writer, sheet_name='MoM Trend Analysis', startrow=start_sku_row, index=False)

            for col_i in range(len(df_ch_mom_exp.columns)):
                ws_mom.set_column(col_i, col_i, 22, cell_fmt)
                ws_mom.write(2, col_i, df_ch_mom_exp.columns[col_i], header_fmt)
            for col_i in range(len(df_sku_mom_exp.columns)):
                ws_mom.write(start_sku_row, col_i, df_sku_mom_exp.columns[col_i], header_fmt)

            chart_vol = workbook.add_chart({'type': 'column'})
            max_r = len(df_ch_mom_exp) + 1
            for col_idx_m in range(1, len(df_ch_mom_exp.columns)):
                chart_vol.add_series({
                    'name':       ['MoM Trend Analysis', 2, col_idx_m],
                    'categories': ['MoM Trend Analysis', 3, 0, max_r, 0],
                    'values':     ['MoM Trend Analysis', 3, col_idx_m, max_r, col_idx_m],
                })
            chart_vol.set_title({'name': 'Channel Audit Volume Progression (MoM)'})
            chart_vol.set_size({'width': 650, 'height': 320})
            ws_mom.insert_chart('G3', chart_vol)

        # 5. Helper Function for Channel Tabs
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
            ]).fillna('N/A')

            client_deals = product_listing_deals.get(sheet_name, {})
            sku_records = []
            for orig_col, clean_name in sku_mapping.items():
                col_series = sub_df[orig_col].astype(str)
                has_stock = col_series.str.contains('有貨').sum()
                oos_tag = col_series.str.contains('缺貨').sum()
                no_tag = col_series.str.contains('無貨').sum()
                cov = round((has_stock / tot_visits) * 100, 1) if tot_visits > 0 else 0
                cat = get_product_category(clean_name)
                
                if sheet_name == "All Stores (Overall)":
                    deal_note = "Aggregated Overall"
                    should_include = True
                else:
                    has_deal = match_sku_deal(clean_name, client_deals)
                    if has_stock > 0 and has_deal is True:
                        deal_note = "With Deal"
                        should_include = True
                    elif has_stock > 0 and has_deal is False:
                        deal_note = "Unlisted / Stocked without Deal (No Deal but in stock)"
                        should_include = True
                    elif has_stock == 0 and has_deal is True:
                        deal_note = "With Deal (OOS)"
                        should_include = True
                    else:
                        deal_note = "0% due to No Deal (Not Listed)"
                        should_include = False

                if should_include:
                    sku_records.append({
                        "Category": cat,
                        "Product SKU": clean_name,
                        "Deal Status": deal_note,
                        "有貨有牌仔": has_stock,
                        "缺貨有牌仔": oos_tag,
                        "無貨無牌仔": no_tag,
                        "Coverage (%)": f"{cov}%",
                        "Performance Rating": get_coverage_rating(cov)
                    })
            df_sku_details = pd.DataFrame(sku_records).fillna('N/A')

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

            if sheet_name != "All Stores (Overall)" and len(df_sku_details) > 0:
                chart_shop = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
                max_sku_row = len(df_sku_details) + 7
                chart_shop.add_series({
                    'name':       [sheet_name, 7, 3],
                    'categories': [sheet_name, 8, 1, min(max_sku_row, 30), 1],
                    'values':     [sheet_name, 8, 3, min(max_sku_row, 30), 3],
                    'fill':       {'color': '#2ca02c'}
                })
                chart_shop.add_series({
                    'name':       [sheet_name, 7, 4],
                    'categories': [sheet_name, 8, 1, min(max_sku_row, 30), 1],
                    'values':     [sheet_name, 8, 4, min(max_sku_row, 30), 4],
                    'fill':       {'color': '#d62728'}
                })
                chart_shop.add_series({
                    'name':       [sheet_name, 7, 5],
                    'categories': [sheet_name, 8, 1, min(max_sku_row, 30), 1],
                    'values':     [sheet_name, 8, 5, min(max_sku_row, 30), 5],
                    'fill':       {'color': '#a6a6a6'}
                })
                chart_shop.set_title({'name': f'{sheet_name} - Listed Products Availability'})
                chart_shop.set_size({'width': 750, 'height': 380})
                ws.insert_chart('J8', chart_shop)

        export_detailed_channel_sheet(df, 'All Stores (Overall)')
        for ch_label in valid_channels:
            sub_df = df[df['店鋪_clean'].str.contains(ch_label, regex=False, na=False)]
            sheet_title = str(ch_label).replace(':', '').replace('/', '-')[:30]
            export_detailed_channel_sheet(sub_df, sheet_title)

    st.download_button(
        label=f"🟢 Download Complete Formatted Excel Report ({active_m_en})",
        data=output.getvalue(),
        file_name=f"Market_Visit_Summary_{active_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )