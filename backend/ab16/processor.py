import re
import pandas as pd
from ..processors.base import BaseProcessor
from typing import List, Dict, Any
from ..core.excel_utils import get_cell_val
from ..core.fuzzy_utils import fuzzy_match_name, robust_clean
from ..config import COMMON_CELLS, MASTER_COUNTRIES

class AB16Processor(BaseProcessor):
    def __init__(self, category_type: str, master_manufacturers: Dict[str, str] = None):
        # category_type could be 'AB16_TERM' or 'AB16_CLEAT'
        self.category_type = category_type
        if master_manufacturers is None:
            from ..core.data_loader import load_master_manufacturers
            self.master_manufacturers = load_master_manufacturers()
        else:
            self.master_manufacturers = master_manufacturers

    def get_category(self) -> str:
        return self.category_type

    def process_sheet(self, ws, filename: str, master_equip: Dict[str, str], **kwargs) -> Dict[str, Any]:
        """
        Processes an AB16 sheet. For AB16_CLEAT, performs full specs extraction and validation.
        For AB16_TERM, falls back to the original manufacturer extraction stub.
        """
        if self.category_type == "AB16_CLEAT":
            return self._process_cleat_sheet(ws, filename, master_equip, **kwargs)
        else:
            return self._process_term_sheet(ws, filename, master_equip, **kwargs)

    def _process_cleat_sheet(self, ws, filename: str, master_equip: Dict[str, str], **kwargs) -> Dict[str, Any]:
        bidder_raw = get_cell_val(ws, "C6")
        if not bidder_raw:
            return {}

        proc_ref = get_cell_val(ws, "E5")
        sched_no = get_cell_val(ws, "K5")
        item_no = str(get_cell_val(ws, "J6")).upper()

        # 1. Check if B15 contains "For XLPE Cable Suitability"
        b15_val = get_cell_val(ws, "B15")
        if not b15_val or "For XLPE Cable Suitability" not in str(b15_val):
            return {
                "BidderRaw": bidder_raw,
                "ProcRef": proc_ref,
                "Schedule": sched_no,
                "ItemNo": item_no or "Unknown",
                "Status": "INVALID",
                "Issues": ["Wrong version (B15 must contain 'For XLPE Cable Suitability')"],
                "Equipment": {}
            }

        # 2. Extract technical values
        mfr_raw = get_cell_val(ws, "G11")
        country_raw = get_cell_val(ws, "G12")
        type_model = get_cell_val(ws, "G14")
        min_od = get_cell_val(ws, "I16")
        max_od = get_cell_val(ws, "N16")
        applied_std = get_cell_val(ws, "G18")
        supp_std = get_cell_val(ws, "G20")

        g22 = get_cell_val(ws, "G22")
        j23 = get_cell_val(ws, "J23")
        g24 = get_cell_val(ws, "G24")
        j25 = get_cell_val(ws, "J25")
        j26 = get_cell_val(ws, "J26")

        spacing_s = get_cell_val(ws, "G28")
        peak_current = get_cell_val(ws, "G29")
        spacing_d = get_cell_val(ws, "G30")
        formation = get_cell_val(ws, "G32")

        # 3. Perform Validations
        # A. Applied standard validation
        std_str = str(applied_std or "").strip()
        req_std = "conforming with EGAT requirement (IEC 61914)"
        std_ok = ("IEC 61914" in std_str) and ("EGAT" in std_str or "requirement" in std_str.lower())
        if not std_ok and std_str.lower() == req_std.lower():
            std_ok = True

        # B. Formation validation
        form_str = str(formation or "").strip()
        formation_ok = form_str.upper() in ["TREFOIL", "FLAT"]

        # C. Material Classification validation
        has_metallic = bool(g22 and str(g22).strip())
        has_composite = bool(g24 and str(g24).strip())
        mat_ok = True
        material_class = ""
        material_detail = ""

        if has_metallic and has_composite:
            material_class = "Both Metallic & Composite"
            material_detail = f"Metallic: {j23} / Composite: {j25} {j26}"
            mat_ok = False
        elif has_metallic:
            material_class = "Metallic"
            material_detail = str(j23 or "").strip()
            if not material_detail:
                mat_ok = False
        elif has_composite:
            material_class = "Composite"
            parts = [p for p in [j25, j26] if p and str(p).strip()]
            material_detail = " / ".join(str(p).strip() for p in parts)
            if not material_detail:
                mat_ok = False
        else:
            material_class = "None"
            mat_ok = False

        # D. AB17 Cable Compatibility Audit
        ab17_data = kwargs.get("ab17_cable_data", {})
        print(f"[DEBUG] AB16 Audit: Available AB17 keys: {list(ab17_data.keys())}")
        cable_audit_msg = ""
        cable_audit_ok = True
        
        # Look for description in price schedule
        item_desc = master_equip.get(item_no, "").upper()
        print(f"[DEBUG] AB16 Audit: Item {item_no} Description: {item_desc}")
        
        # Keyword Detection: TREFOIL or FLAT in description
        reasons = []
        desc_formation = None
        if "TREFOIL" in item_desc:
            desc_formation = "TREFOIL"
        elif "FLAT" in item_desc:
            desc_formation = "FLAT"
            
        if desc_formation and desc_formation != form_str.upper():
            reasons.append(f"Formation mismatch: Description says {desc_formation}, but PD says {form_str.upper()}")

        if "AB17" in item_desc:
            # Extract AB17 ID from description or item_no if possible
            # Assuming AB17 ID format is XAB17-xxxx
            match = re.search(r"\d+AB17-\d+", item_desc)
            if match:
                target_ab17_id = match.group(0)
                print(f"[DEBUG] AB16 Audit: Target AB17 ID: {target_ab17_id}")
                if target_ab17_id in ab17_data:
                    c_data = ab17_data[target_ab17_id]
                    c_min, c_max = c_data["min"], c_data["max"]
                    print(f"[DEBUG] AB16 Audit: Cable dimensions {c_min}-{c_max}")
                    if not (min_od <= c_min and max_od >= c_max):
                        cable_audit_ok = False
                        cable_audit_msg = f"Cleat OD ({min_od}-{max_od}) incompatible with Cable {target_ab17_id} ({c_min}-{c_max})"
                else:
                    cable_audit_ok = False
                    cable_audit_msg = f"AB17 Item {target_ab17_id} not found in proposal documents."
        
        # Gather audit results
        if not std_ok:
            reasons.append("Standard mismatch")
        if not formation_ok:
            reasons.append(f"Invalid Formation: '{form_str}' (must be Trefoil or Flat)")
        if not mat_ok:
            reasons.append("Invalid material classification")
        if not cable_audit_ok:
            reasons.append(cable_audit_msg)

        row_result = "Pass"
        row_comment = "Pass"
        if reasons:
            row_result = "Fail"
            row_comment = " & ".join(reasons)

        status = "VALID"
        issues = []
        if reasons:
            status = "INVALID"
            issues.extend(reasons)

        # Normalize manufacturer and country
        m_final = fuzzy_match_name(mfr_raw, self.master_manufacturers, threshold=90) if mfr_raw else ""
        c_clean = robust_clean(country_raw) if country_raw else ""
        c_final = MASTER_COUNTRIES.get(c_clean, country_raw) if country_raw else ""

        # Map to equip_order
        equip_order = kwargs.get("equip_order", [])
        eq_key = next((k for k in equip_order if item_no in k), item_no or "Unknown")

        equipment_rows = [{
            "manufacturer": m_final,
            "country": c_final,
            "standard": f"{std_str}\n{supp_std}".strip() if supp_std else std_str,
            "type_model": type_model,
            "od_range": f"{min_od} - {max_od}" if min_od or max_od else "",
            "material_class": material_class,
            "material_detail": material_detail,
            "spacing_s": spacing_s,
            "spacing_d": spacing_d,
            "peak_current": peak_current,
            "formation": formation,
            "result": row_result,
            "comment": row_comment,
            "source": f"{filename} > {ws.title}"
        }]

        return {
            "BidderRaw": bidder_raw,
            "ProcRef": proc_ref,
            "Schedule": sched_no,
            "ItemNo": item_no,
            "Status": status,
            "Issues": issues,
            "Equipment": {
                eq_key: equipment_rows
            }
        }

    def _process_term_sheet(self, ws, filename: str, master_equip: Dict[str, str], **kwargs) -> Dict[str, Any]:
        import pandas as pd
        from .term_logic import extract_fields_from_pd, check_qwerty_logic, _numbers_from_text
        from ..shared.proposal_helpers import extract_item_id_from_pd
        
        bidder_raw = ""
        for cell in ["C7", "C6", "C11", "C5", "C8", "C9", "C10"]:
            val = str(get_cell_val(ws, cell) or "").strip()
            if val and "MANUFACTURER" not in val.upper() and "BIDDER" not in val.upper() and len(val) > 2:
                bidder_raw = val
                break
                
        if not bidder_raw: return {}

        df_pd = pd.DataFrame(ws.values)
        item_no = str(get_cell_val(ws, "J7") or "").strip() or extract_item_id_from_pd(df_pd) or "Unknown"
        
        proc_ref = get_cell_val(ws, "E6") or get_cell_val(ws, "E5") or get_cell_val(ws, "F8")
        sched_no = get_cell_val(ws, "K6") or get_cell_val(ws, "K5") or get_cell_val(ws, "L8") or get_cell_val(ws, "L6") or get_cell_val(ws, "M6")

        # Extract fields
        pd_fields = extract_fields_from_pd(df_pd)

        # Get PS Description
        ps_description = master_equip.get(item_no, "")

        # Map to equip_order
        equip_order = kwargs.get("equip_order", [])
        eq_key = next((k for k in equip_order if item_no in k), item_no)

        # Find AB17 Cable bounds
        ab17_data = kwargs.get("ab17_cable_data", {})
        ab17_min, ab17_max = None, None
        ref_fields = {}

        if "AB17" in ps_description.upper():
            match = re.search(r"\d+AB17-\d+", ps_description)
            if match:
                target_ab17_id = match.group(0)
                if target_ab17_id in ab17_data:
                    c_data = ab17_data[target_ab17_id]
                    # We might have missing bounds but have f, g, h
                    c_min_val = c_data.get("min")
                    c_max_val = c_data.get("max")
                    if c_min_val and c_max_val:
                        mins = _numbers_from_text(str(c_min_val))
                        maxs = _numbers_from_text(str(c_max_val))
                        if mins and maxs:
                            ab17_min = mins[0]
                            ab17_max = maxs[0]
                    # Pass f, g, h
                    ref_fields = {
                        "f_cond_od": c_data.get("f_cond_od", ""),
                        "g_cond_screen": c_data.get("g_cond_screen", ""),
                        "h_insulation": c_data.get("h_insulation", "")
                    }

        # Run QWERTY logic
        final_status, final_comment = check_qwerty_logic(
            pd_fields, ps_description, ab17_min=ab17_min, ab17_max=ab17_max, ref_fields=ref_fields
        )

        status = "VALID" if "ผ่านเงื่อนไขเบื้องต้น" in final_status else "INVALID"
        
        issues = []
        if status == "INVALID":
            # Extract bullet points from final_status
            lines = final_status.split('\n')
            for line in lines:
                line = line.strip()
                if line.startswith(("a.", "b.", "c.", "d.", "e.", "f.", "g.", "h.", "i.", "j.", "k.")):
                    issues.append(line)

        # Use fuzzy match for manufacturer/country if needed, or just what we parsed
        mfr_raw = pd_fields.get("Manufacturer", "")
        country_raw = pd_fields.get("Country of origin", "")
        m_final = fuzzy_match_name(mfr_raw, self.master_manufacturers, threshold=90) if mfr_raw else ""
        c_clean = robust_clean(country_raw) if country_raw else ""
        c_final = MASTER_COUNTRIES.get(c_clean, country_raw) if country_raw else ""

        equipment_rows = [{
            "manufacturer": m_final or mfr_raw,
            "country": c_final or country_raw,
            "standard": pd_fields.get("Applied Standard", ""),
            "type_model": pd_fields.get("Type / Model", ""),
            "od_range": f"{pd_fields.get('Insulation Diameter Min', '')} - {pd_fields.get('Insulation Diameter Max', '')}".strip(" -"),
            "result": "Pass" if status == "VALID" else "Fail",
            "comment": final_status + "\n\n" + final_comment,
            "source": f"{filename} > {ws.title}"
        }]

        return {
            "BidderRaw": bidder_raw,
            "ProcRef": proc_ref,
            "Schedule": sched_no,
            "ItemNo": item_no,
            "Status": status,
            "Issues": issues,
            "Equipment": {
                eq_key: equipment_rows
            }
        }

