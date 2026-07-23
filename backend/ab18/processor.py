import re
from typing import List, Dict, Any, Tuple
from ..processors.base import BaseProcessor
from ..core.excel_utils import get_cell_val
from ..core.fuzzy_utils import fuzzy_match_name, robust_clean
from ..config import COMMON_CELLS, MASTER_COUNTRIES
from thefuzz import fuzz, process as fuzz_process

# AB18 Specific Configuration
EQUIP_CELLS = {
    "Power Cable": {"m": "H15", "c": "H16", "sf": "H18", "ss": "H20", "req_sf": "conforming with EGAT requirement (TIS11-2559)"},
    "Control Cable": {"m": "H15", "c": "H16", "sf": "N18", "ss": "N20", "req_sf": "conforming with EGAT requirement (TIS838-2531)"},
    "Twisted Pair Control Cable": {"m": "H15", "c": "H16", "sf": "R18", "ss": "R20", "req_sf": "conforming with EGAT requirement (TIS838-2531)"},
    "THW Lighting Cable": {"m": "H23", "c": "H24", "sf": "H25", "ss": "H27", "req_sf": "conforming with EGAT requirement (TIS11-2553)"},
    "NYY Lighting Cable": {"m": "P23", "c": "P24", "sf": "P25", "ss": "P27", "req_sf": "conforming with EGAT requirement (TIS11-2559)"},
    "Annealed Copper Ground Wire": {"m": "H30", "c": "H31", "sf": "H32", "ss": "H34", "req_sf": "conforming with EGAT requirement (ASTM B3, ASTM B8)"},
    "Aluminum Conductor": {"m": "H44", "c": "H45", "sf": "H46", "ss": "H48", "req_sf": "conforming with EGAT requirement (ASTM B231, ASTM B232)"},
    "Overhead Ground Wire": {"m": "H37", "c": "H38", "sf": "H39", "ss": "H41", "req_sf": "conforming with EGAT requirement (ASTM A363)"}
}

def expand_item_ranges(item_no_raw: str) -> List[str]:
    if not item_no_raw:
        return []
    item_no_raw = item_no_raw.replace('\n', ',')
    chunks = [c.strip() for c in re.split(r'[,/]| AND |&', item_no_raw.upper()) if c.strip()]
    items = []
    for chunk in chunks:
        m = re.search(r'(\d+AB\d+)-(\d+)\s*(?:THRU|TO|-)\s*(?:(\d+AB\d+)-)?(\d+)', chunk, re.I)
        if m:
            prefix = m.group(1).upper()
            start_num = int(m.group(2))
            end_prefix = m.group(3)
            end_num = int(m.group(4))
            if end_prefix and end_prefix.upper() != prefix:
                items.append(chunk)
            else:
                for n in range(start_num, end_num + 1):
                    items.append(f"{prefix}-{n}")
        else:
            items.append(chunk)
    return items

class AB18Processor(BaseProcessor):
    def __init__(self, master_manufacturers: Dict[str, str], egat_pools: Tuple[List[str], List[str]]):
        self.master_manufacturers = master_manufacturers
        self.egat_general_mfrs, self.egat_overhead_mfrs = egat_pools

    def get_category(self) -> str:
        return "AB18"

    def split_paired_data(self, m_cell: str, c_cell: str, sf_cell: str, ss_cell: str, 
                          filename: str, sheetname: str, req_standard: str = "", 
                          group_id: str = "") -> List[Dict]:
        def split_val(v):
            if not v: return []
            for sep in ["/", " AND ", " OR ", "&", ";", "\n"]: v = v.replace(sep, "|")
            return [p.strip() for p in v.split("|") if p.strip()]

        m_list = split_val(m_cell)
        c_list = split_val(c_cell)
        sf_list = split_val(sf_cell)
        ss_list = split_val(ss_cell)
        
        max_len = max(len(m_list), len(c_list), len(sf_list), len(ss_list))
        if max_len == 0: return []
        
        if len(m_list) > 1:
            if len(c_list) == 1: c_list = c_list * len(m_list)
            if len(sf_list) == 1: sf_list = sf_list * len(m_list)
            if len(ss_list) == 1: ss_list = ss_list * len(m_list)

        m_list += [""] * (max_len - len(m_list))
        c_list += [""] * (max_len - len(c_list))
        sf_list += [""] * (max_len - len(sf_list))
        ss_list += [""] * (max_len - len(ss_list))
        
        results = []
        for m, c, sf, ss in zip(m_list, c_list, sf_list, ss_list):
            m_final = fuzzy_match_name(m, self.master_manufacturers)
            c_clean = robust_clean(c)
            # Ensure we keep the cleaned name if the master mapping fails
            c_final = MASTER_COUNTRIES.get(c_clean, c_clean if c_clean else c)
            
            std_match = (sf == req_standard)
            s_combined = f"{sf}\n{ss}".strip() if ss else sf
            
            is_ovh = (group_id == "Overhead Ground Wire")
            approved_pool = self.egat_overhead_mfrs if is_ovh else self.egat_general_mfrs
            
            m_approved = False
            if m_final and approved_pool:
                m_upper = m_final.upper()
                pool_upper = [p.upper() for p in approved_pool]
                if m_upper in pool_upper:
                    m_approved = True
                else:
                    match, score = fuzz_process.extractOne(m_upper, pool_upper)
                    if score >= 85: m_approved = True
            
            result, comment = "", ""
            if sf or ss or m_final:
                if std_match and m_approved:
                    result = "Pass"
                else:
                    result = "Fail"
                    reasons = []
                    if not std_match: reasons.append("Standard mismatch")
                    if not m_approved: reasons.append("Manufacturer not in EGAT Pool")
                    comment = " & ".join(reasons)
            
            results.append({
                "manufacturer": m_final,
                "country": c_final,
                "standard": s_combined,
                "result": result,
                "comment": comment,
                "source": f"{filename} > {sheetname}"
            })
        return results

    def process_sheet(self, ws, filename: str, master_equip: Dict[str, str], **kwargs) -> Dict[str, Any]:
        bidder_raw = get_cell_val(ws, COMMON_CELLS["bidder"])
        if not bidder_raw: return {}
        
        status, issues = "VALID", []
        
        # Robust filtering for non-equipment sheets/items
        unwanted_terms = [
            "TOTAL PRICE", "SUMMARY", "TRANSPORTATION", 
            "CONSTRUCTION AND INSTALLATION", "LOCAL TRANSPORTATION",
            "TOTAL PRICE FOR SCHEDULE", "GRAND TOTAL"
        ]
        
        ws_title = ws.title.upper()
        item_no_raw = get_cell_val(ws, COMMON_CELLS["item_no"])
        
        if any(term in ws_title for term in unwanted_terms) or \
           any(term in item_no_raw.upper() for term in unwanted_terms):
            return {}

        sched = get_cell_val(ws, COMMON_CELLS["schedule_no"]).upper()
        if not re.match(r'^\d+AB18$', sched):
            status = "INCORRECT PD"
            issues.append(f"Invalid Schedule at L8: {sched}")
        
        item_no_raw = get_cell_val(ws, COMMON_CELLS["item_no"])
        clean_items = expand_item_ranges(item_no_raw)
        clean_items = [it for it in clean_items if not any(term in it for term in ["TRANSPORTATION", "TOTAL PRICE", "SUMMARY"])]

        if not clean_items:
            return {}

        equip_data = {}
        equip_order = kwargs.get("equip_order", [])

        for group_id, cells in EQUIP_CELLS.items():
            m = get_cell_val(ws, cells["m"])
            c = get_cell_val(ws, cells["c"])
            sf = get_cell_val(ws, cells["sf"])
            ss = get_cell_val(ws, cells["ss"])
            
            is_special = group_id in ["Power Cable", "Control Cable", "Twisted Pair Control Cable"]
            has_standard = bool(sf or ss)
            has_any_data = any([m, c, sf, ss])
            
            # Check if item is in price schedule
            if "THW" in group_id or "NYY" in group_id:
                is_in_schedule = any(fuzz.token_set_ratio(group_id.upper(), v.upper()) >= 90 for v in master_equip.values())
            elif group_id == "Control Cable":
                is_in_schedule = any(group_id.upper() in v.upper() and "TWISTED PAIR" not in v.upper() for v in master_equip.values())
            else:
                is_in_schedule = any(group_id.upper() in v.upper() for v in master_equip.values())

            process_data = False
            
            if is_special:
                if is_in_schedule:
                    if has_standard:
                        if sf != cells["req_sf"]:
                            status = "INVALID"
                            issues.append(f"Standard mismatch for {group_id}")
                        process_data = True
                else:
                    if has_standard:
                        status = "INVALID"
                        issues.append(f"Unauthorized item filled: {group_id}")
                        process_data = True
            else:
                if has_any_data:
                    process_data = True
                    if not is_in_schedule:
                        status = "INVALID"
                        issues.append(f"Unauthorized item filled: {group_id}")
                    elif sf != cells["req_sf"]:
                        status = "INVALID"
                        issues.append(f"Standard mismatch for {group_id}")
            
            if process_data:
                matching_keys = []
                if "THW" in group_id or "NYY" in group_id:
                    matching_keys = [k for k in equip_order if fuzz.token_set_ratio(group_id.upper(), k.upper()) >= 90]
                elif group_id == "Control Cable":
                    matching_keys = [k for k in equip_order if group_id.upper() in k.upper() and "TWISTED PAIR" not in k.upper()]
                else:
                    matching_keys = [k for k in equip_order if group_id.upper() in k.upper()]
                
                # Assign to all matching equipment keys that belong to clean_items
                filtered_keys = [k for k in matching_keys if k.split()[0] in clean_items]
                if not filtered_keys:
                    filtered_keys = matching_keys
                
                for eq_key in filtered_keys:
                    equip_data[eq_key] = self.split_paired_data(m, c, sf, ss, filename, ws.title, req_standard=cells["req_sf"], group_id=group_id)

        return {
            "BidderRaw": bidder_raw,
            "ProcRef": get_cell_val(ws, COMMON_CELLS["procurement_ref"]),
            "Schedule": sched,
            "ItemNo": ", ".join(clean_items),
            "Status": status,
            "Issues": issues,
            "Equipment": equip_data
        }

