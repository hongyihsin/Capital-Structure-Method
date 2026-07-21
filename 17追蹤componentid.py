import pandas as pd
from pathlib import Path
import re

# ==========================================
# [CONFIG] 使用者參數設定區
# ==========================================
# 請在此輸入想要追蹤的 componentid (支援數字或字串)
TARGET_COMPONENT_ID = "903139141"

# 檔案與資料夾路徑設定
DATA_DIR = Path(r"C:\ntutdata\debt")
SIC_EXCEL_PATH = Path(r"C:\Users\olivi\Desktop\acer\ntut\1Seminar\data\GVKEY_debt structure_with SIC.xlsx")

# ==========================================
# Helper Functions (工具函數)
# ==========================================
def extract_column_name(col):
    match = re.match(r"\((.*?)\)", str(col))
    return match.group(1).strip().lower() if match else str(col).strip().lower()

def clean_to_string(val):
    if pd.isna(val) or str(val).strip().lower() in ['nan', 'null', '']:
        return None
    return str(val).strip().split('.')[0]

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

target_id_str = clean_to_string(TARGET_COMPONENT_ID)

print("="*80)
print(f" [DEBT TRAJECTORY TRACKER] Searching ComponentID: [{target_id_str}]")
print("="*80)

# ==========================================
# 1. 載入 GVKEY 對照表
# ==========================================
company_to_gvkey = {}
try:
    sic_df = pd.read_excel(SIC_EXCEL_PATH)
    sic_df.columns = [extract_column_name(c) for c in sic_df.columns]
    gvkey_col = [c for c in sic_df.columns if "gvkey" in c]
    companyid_col = [c for c in sic_df.columns if "companyid" in c]
    if gvkey_col and companyid_col:
        sic_df["companyid"] = sic_df[companyid_col[0]].apply(clean_to_string)
        sic_df["gvkey"] = sic_df[gvkey_col[0]].apply(clean_to_string)
        company_to_gvkey = dict(zip(sic_df["companyid"], sic_df["gvkey"]))
except Exception as e:
    print(f"[WARN] GVKEY Mapping failed to load: {e}")

# ==========================================
# 2. 跨檔案搜尋指定的 ComponentID
# ==========================================
files = list(DATA_DIR.glob("*.xlsx")) + list(DATA_DIR.glob("*.csv"))
history_records = []

print(f"[INFO] Scanning files in {DATA_DIR}...")

for f in files:
    if "analysis" in f.name.lower() or "report" in f.name.lower(): continue
    try:
        df = pd.read_excel(f) if f.suffix == ".xlsx" else pd.read_csv(f, encoding="cp950", errors="ignore")
        df.columns = [extract_column_name(c) for c in df.columns]
        
        if "componentid" not in df.columns: continue
        
        df["comp_id_clean"] = df["componentid"].apply(clean_to_string)
        
        matched = df[df["comp_id_clean"] == target_id_str].copy()
        
        if not matched.empty:
            matched["_Source_File"] = f.name
            history_records.append(matched)
            
    except Exception as e:
        pass

# ==========================================
# 3. 整理時間軸與金額變動軌跡
# ==========================================
if history_records:
    traj_df = pd.concat(history_records, ignore_index=True)
    
    traj_df["companyid"] = traj_df["companyid"].apply(clean_to_string)
    traj_df["gvkey"] = traj_df["companyid"].map(company_to_gvkey).fillna(traj_df["companyid"])
    traj_df["Actual_Amount"] = traj_df.apply(lambda r: calc_actual_value(r.get("dataitemvalue"), r.get("unittypeid")), axis=1)
    traj_df["Debt_Category"] = traj_df.apply(lambda r: map_debt_category(r.get("capitalstructuresubtypeid"), r.get("leveltypeid")), axis=1)
    
    traj_df["Report_Date"] = pd.to_datetime(traj_df["periodenddate"], errors='coerce')
    traj_df = traj_df.dropna(subset=["Report_Date"]).sort_values(by="Report_Date").reset_index(drop=True)
    
    traj_df["Amount_Change"] = traj_df["Actual_Amount"].diff().fillna(0)
    traj_df["Pct_Change(%)"] = (traj_df["Actual_Amount"].pct_change() * 100).fillna(0)
    
    first_row = traj_df.iloc[0]
    last_row = traj_df.iloc[-1]
    
    print("\n" + "="*80)
    print(f" [DEBT PROFILE SUMMARY]")
    print(f"  * Component ID : {target_id_str}")
    print(f"  * Company ID   : {first_row['companyid']} (GVKEY: {first_row['gvkey']})")
    print(f"  * Debt Category: {first_row['Debt_Category']}")
    print(f"  * Description  : {first_row.get('descriptiontext', 'N/A')}")
    print(f"  * Lifespan     : {first_row['Report_Date'].strftime('%Y-%m-%d')} to {last_row['Report_Date'].strftime('%Y-%m-%d')} ({len(traj_df)} records found)")
    print("="*80)
    
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 1000)
    pd.set_option('display.float_format', lambda x: '%.2f' % x)
    
    display_cols = [
        "Report_Date", 
        "dataitemvalue", 
        "unittypeid", 
        "Actual_Amount", 
        "Amount_Change", 
        "Pct_Change(%)", 
        "_Source_File"
    ]
    present_cols = [c for c in display_cols if c in traj_df.columns]
    
    print("\n[ Historical Debt Trajectory Timeline ]")
    print(traj_df[present_cols].to_string(index=False))
    print("="*80)
    
    print("\n[ Trajectory Status Analysis ]")
    if len(traj_df) == 1:
        print("  * Status: Single Record (The debt appears only once in the dataset).")
    else:
        init_val = traj_df.iloc[0]["Actual_Amount"]
        final_val = traj_df.iloc[-1]["Actual_Amount"]
        if final_val == 0:
            print("  * Status: Paid Off (The debt amount has decreased to zero in the latest record).")
        elif final_val < init_val:
            print(f"  * Status: Amortizing (Decreased from {init_val:,.0f} to {final_val:,.0f}, -{(1-final_val/init_val)*100:.1f}%).")
        elif final_val > init_val:
            print(f"  * Status: Increasing (Increased from {init_val:,.0f} to {final_val:,.0f}, +{(final_val/init_val-1)*100:.1f}%).")
        else:
            print("  * Status: Constant (The debt amount remained unchanged during the tracked period).")
    print("="*80)

else:
    print(f"\n[ERROR] ComponentID [{target_id_str}] was not found in any dataset files.")