import pandas as pd
from pathlib import Path
import re

# ==========================================
# ⚙️ [CONFIG] 使用者參數設定區（在此修改即可）
# ==========================================
# 1. 選擇要檢查的年份 (例如 1999, 2008, 2011)；若設為 None 則顯示全年度摘要
TARGET_YEAR = 1999 

# 2. 選擇要檢查的債務類別 (例如 "SBN", "CP", "TL", "Others"；設為 None 則檢查該年所有類別)
TARGET_CATEGORY = "SBN"

# 3. 檔案與資料夾路徑設定
DATA_DIR = Path(r"C:\ntutdata\debt")
SIC_EXCEL_PATH = Path(r"C:\Users\olivi\Desktop\acer\ntut\1Seminar\data\GVKEY_debt structure_with SIC.xlsx")

# ==========================================
# 🛠️ Helper Functions (工具函數)
# ==========================================
def extract_column_name(col):
    match = re.match(r"\((.*?)\)", str(col))
    return match.group(1).strip().lower() if match else str(col).strip().lower()

def clean_to_string(val):
    if pd.isna(val) or str(val).strip().lower() in ['nan', 'null', '']:
        return None
    return str(val).strip().split('.')[0]

def parse_year(date_val):
    if pd.isna(date_val): return None
    txt = str(date_val).strip()
    m = re.search(r"(19\d{2}|20\d{2})", txt)
    if m: return int(m.group())
    try:
        return pd.to_datetime(txt, errors='coerce').year
    except:
        return None

def calc_actual_value(value, unit):
    try: value = float(value)
    except: value = 0.0
    try: unit = int(unit)
    except: unit = 0
    
    if unit == 1: return value * 1000
    elif unit == 2: return value * 1000000
    return value

def map_debt_category(subtype, level):
    try: subtype = int(subtype)
    except: subtype = -1
    try: level = int(level)
    except: level = -1

    if subtype == 1: return "CP"
    elif subtype == 2: return "DC"
    elif subtype == 3: return "TL"
    elif subtype == 5: return "CL"
    elif subtype in [6, 7, 9]: return "Others"
    elif subtype == 4:
        if level == 1: return "SBN"
        elif level in [2, 3, 4, 5, 6, 7]: return "SUB"
    return "Others"

# ==========================================
# 1. 載入白名單 (Whitelist)
# ==========================================
print(f"[INFO] Loading SIC industry whitelist...")
sic_df = pd.read_excel(SIC_EXCEL_PATH)
sic_df.columns = [extract_column_name(c) for c in sic_df.columns]

gvkey_col = [c for c in sic_df.columns if "gvkey" in c]
companyid_col = [c for c in sic_df.columns if "companyid" in c]
sic_col = [c for c in sic_df.columns if "sic" in c and c != "sic_int"]

sic_df["companyid"] = sic_df[companyid_col[0]].apply(clean_to_string)
sic_df["sic_int"] = pd.to_numeric(sic_df[sic_col[0]], errors="coerce").fillna(-1).astype(int)

filtered_sic = sic_df[
    ~(((sic_df["sic_int"] >= 4900) & (sic_df["sic_int"] <= 4949)) |
      ((sic_df["sic_int"] >= 6000) & (sic_df["sic_int"] <= 6999)))
]
valid_companies = set(filtered_sic["companyid"].dropna().unique())
company_to_gvkey = dict(zip(filtered_sic["companyid"], filtered_sic[gvkey_col[0]].apply(clean_to_string))) if gvkey_col else {}

# ==========================================
# 2. 檢索與讀取資料
# ==========================================
print(f"[INFO] Scanning data directory: {DATA_DIR}")
files = list(DATA_DIR.glob("*.xlsx")) + list(DATA_DIR.glob("*.csv"))
matched_records = []

for f in files:
    if "analysis" in f.name.lower() or "report" in f.name.lower(): continue
    try:
        df = pd.read_excel(f) if f.suffix == ".xlsx" else pd.read_csv(f, encoding="cp950", errors="ignore")
        df.columns = [extract_column_name(c) for c in df.columns]
        
        if "companyid" not in df.columns or "periodenddate" not in df.columns: continue
        
        df["companyid"] = df["companyid"].apply(clean_to_string)
        df = df[df["companyid"].isin(valid_companies)].copy()
        
        if "issuedcurrencyid" in df.columns:
            df["issuedcurrencyid"] = pd.to_numeric(df["issuedcurrencyid"], errors="coerce")
            df = df[df["issuedcurrencyid"] == 160].copy()
            
        df["Report_Year"] = df["periodenddate"].apply(parse_year)
        
        # 篩選指定年份
        if TARGET_YEAR is not None:
            df = df[df["Report_Year"] == TARGET_YEAR].copy()
            
        if df.empty: continue
        
        df["Debt_Category"] = df.apply(lambda r: map_debt_category(r.get("capitalstructuresubtypeid"), r.get("leveltypeid")), axis=1)
        
        # 篩選指定債務類別
        if TARGET_CATEGORY is not None:
            df = df[df["Debt_Category"] == TARGET_CATEGORY].copy()
            
        if not df.empty:
            df["_Source_File"] = f.name
            matched_records.append(df)
            
    except Exception as e:
        pass

# ==========================================
# 3. 輸出診斷分析結果
# ==========================================
print("\n" + "="*80)
print(f" [DIAGNOSTIC REPORT] Target Year: {TARGET_YEAR if TARGET_YEAR else 'ALL'} | Category: {TARGET_CATEGORY if TARGET_CATEGORY else 'ALL'}")
print("="*80)

if matched_records:
    result_df = pd.concat(matched_records, ignore_index=True)
    result_df["Actual_Amount"] = result_df.apply(lambda r: calc_actual_value(r.get("dataitemvalue"), r.get("unittypeid")), axis=1)
    result_df["gvkey"] = result_df["companyid"].map(company_to_gvkey).fillna(result_df["companyid"])
    
    # 清理 componentid 顯示格式 (避免 .0)
    if "componentid" in result_df.columns:
        result_df["componentid"] = result_df["componentid"].apply(clean_to_string)
    
    # 按金額由大到小排序
    sorted_df = result_df.sort_values(by="Actual_Amount", ascending=False)
    
    # 設定整齊印出格式
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 1000)
    pd.set_option('display.float_format', lambda x: '%.0f' % x)
    
    # 🎯【修改點】已將 componentid 加入輸出欄位清單！
    show_cols = [
        "companyid", 
        "gvkey", 
        "componentid", 
        "periodenddate", 
        "dataitemvalue", 
        "unittypeid", 
        "Debt_Category", 
        "Actual_Amount", 
        "descriptiontext", 
        "_Source_File"
    ]
    present_cols = [c for c in show_cols if c in sorted_df.columns]
    
    print("\n[Top 10 Largest Outlier Records]")
    print(sorted_df[present_cols].head(10).to_string(index=False))
    
    print("\n" + "-"*80)
    print(f"[Summary Statistics]")
    print(f"  * Total Matched Records : {len(sorted_df):,} rows")
    print(f"  * Total Aggregate Amount: {sorted_df['Actual_Amount'].sum():,.0f} USD")
    if len(sorted_df) > 0:
        top_1_ratio = (sorted_df['Actual_Amount'].iloc[0] / sorted_df['Actual_Amount'].sum()) * 100
        print(f"  * Top 1 Share Ratio    : {top_1_ratio:.2f}% of total sum")
    print("="*80)
else:
    print("No matching records found for the specified conditions.")