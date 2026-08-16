def parse_targets_dynamically(df_targets):
    targets_dict = {}
    total_shops_dict = {}
    
    # 1. Dynamically scan first 5 rows to locate actual header row if offset
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
        
    # 2. Dynamically identify channel, target, and total store columns by keyword match
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
            
    # Fallback to relative column positions if keywords are generic
    if ch_col is None and len(df_targets.columns) > 0:
        ch_col = df_targets.columns[0]
    if tg_col is None and len(df_targets.columns) > 1:
        tg_col = df_targets.columns[1]
        
    # 3. Build target mappings dynamically
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