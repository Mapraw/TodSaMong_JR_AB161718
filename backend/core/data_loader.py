import os
import re
import pandas as pd
import openpyxl
from io import BytesIO
from typing import List, Dict, Tuple
from .excel_utils import get_cell_val
from ..config import MASTER_DATA_PATH, MFR_LIST_PATH

def load_master_manufacturers() -> Dict[str, str]:
    """
    Loads manufacturer-alias mapping from the master Excel file.
    """
    master_map = {}
    if not os.path.exists(MASTER_DATA_PATH):
        print(f"Warning: Master Data not found at {MASTER_DATA_PATH}")
        return master_map

    try:
        df_m = pd.read_excel(MASTER_DATA_PATH, sheet_name='Manufacturer')
        raw_map = {}
        current_formal = ""
        for _, row in df_m.iterrows():
            f_name = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            alias = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            if f_name:
                current_formal = f_name
                if current_formal not in raw_map: raw_map[current_formal] = [current_formal]
            if alias and current_formal:
                raw_map[current_formal].append(alias)
        
        for formal, aliases in raw_map.items():
            longest_name = max(aliases, key=len)
            for a in aliases: 
                master_map[a.strip().upper()] = longest_name
    except Exception as e:
        print(f"Error loading master manufacturers: {e}")
    return master_map

def load_egat_mfr_pools() -> Tuple[List[str], List[str]]:
    """
    Loads EGAT's approved manufacturer lists (General and Overhead).
    """
    general_pool = []
    overhead_pool = []
    
    if not os.path.exists(MFR_LIST_PATH):
        print(f"Warning: EGAT Manufacturer list not found at {MFR_LIST_PATH}")
        return general_pool, overhead_pool

    def is_mfr_name(s):
        s = str(s).strip()
        if not s or s.upper().startswith(("NOTE", "TEL", "FAX")): return False
        if re.search(r'\d+/\d+', s): return False
        if any(kw in s.lower() for kw in ["road", "estate", "moo", "km.", "district", "province", "bangkok", "samutprakarn", "pathumthani", "rayong", "chachengsao"]): return False
        return len(s) >= 5

    try:
        gen, ovh = set(), set()
        with pd.ExcelFile(MFR_LIST_PATH) as xls:
            for sheet in xls.sheet_names:
                df_head = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=5)
                if len(df_head) < 2: continue
                
                header_text = str(df_head.iloc[1, 0]).lower()
                is_gen = any(kw in header_text for kw in ["low voltage cable", "ground wire", "aluminum conductor"])
                is_ovh = "overhead ground wire" in header_text
                
                if is_gen or is_ovh:
                    df = pd.read_excel(xls, sheet_name=sheet, header=None)
                    start_idx = -1
                    for idx, v in df.iloc[:, 0].items():
                        if str(v).strip().upper() == "MANUFACTURER":
                            start_idx = idx + 1; break
                    
                    if start_idx != -1:
                        mfrs = [m for m in df.iloc[start_idx:, 0].dropna().astype(str).str.strip().tolist() if is_mfr_name(m)]
                        if is_gen: gen.update(mfrs)
                        if is_ovh: ovh.update(mfrs)
        
        general_pool = sorted(list(gen))
        overhead_pool = sorted(list(ovh))
    except Exception as e:
        print(f"Error loading EGAT Mfrs: {e}")
        
    return general_pool, overhead_pool

def extract_price_schedule_data(content: bytes, category_filter: str = "AB18") -> Tuple[List[str], Dict[str, str]]:
    """
    Extracts equipment IDs and descriptions from a Price Schedule Excel file.
    Supports specific categories (AB16, AB17, AB18) or 'ALL'.
    """
    equip_order = []
    master_equip = {}
    
    try:
        wb = openpyxl.load_workbook(BytesIO(content), data_only=True)
        
        # Identification keywords based on category
        keywords_map = {
            "AB18": ["LOW VOLTAGE CABLE AND CONDUCTOR"],
            "AB17": ["XLPE POWER CABLE"],
            "AB16": ["CABLE TERMINATIONS", "CABLE CLEATS"]
        }
        
        categories_to_process = []
        if category_filter == "ALL":
            categories_to_process = list(keywords_map.keys())
        else:
            categories_to_process = [category_filter]

        for cat in categories_to_process:
            search_terms = keywords_map.get(cat, ["LOW VOLTAGE CABLE AND CONDUCTOR"])
            
            target_sheet = None
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                i4_val = get_cell_val(ws, "I4").upper()
                if any(term in i4_val for term in search_terms):
                    target_sheet = sheet_name
                    break
            
            if target_sheet:
                df = pd.read_excel(BytesIO(content), sheet_name=target_sheet)
                # Assuming similar layout: Col 8 = ID, Col 9 = Description, starting at row 15 (idx 14)
                for i in range(14, len(df)):
                    if len(df.columns) < 10: break
                    item_id = str(df.iloc[i, 8]).strip() if pd.notna(df.iloc[i, 8]) else ""
                    item_desc = str(df.iloc[i, 9]).strip() if pd.notna(df.iloc[i, 9]) else ""
                    
                    # Match by the specific category (AB16, AB17, or AB18)
                    if item_id and cat in item_id.upper():
                        desc_upper = item_desc.upper()
                        # Exclude specific non-equipment terms
                        exclude_terms = [
                            "COST OF LOCAL TRANSPORTATION", 
                            "TOTAL PRICE", 
                            "SUMMARY", 
                            "CONSTRUCTION AND INSTALLATION",
                            "TOTAL PRICE FOR SCHEDULE"
                        ]
                        if any(term in desc_upper for term in exclude_terms):
                            continue
                        
                        item_desc = item_desc.replace("as per Specification attached", "").strip()
                        full_name = f"{item_id} {item_desc}"
                        if full_name not in equip_order:
                            equip_order.append(full_name)
                            master_equip[item_id] = full_name
                
    except Exception as e:
        print(f"Error extracting price schedule ({category_filter}): {e}")
        
    return equip_order, master_equip
