import openpyxl

def get_cell_val(ws, addr):
    """
    Safely retrieves and normalizes a cell value from an openpyxl worksheet.
    Treats common 'blank' symbols like '-', '.', 'N/A' as empty strings.
    """
    try:
        v = ws[addr].value
        if v is None: return ""
        s = str(v).strip()
        # Normalization: Treat common placeholder symbols as blank
        if s in ["-", ".", "N/A", "n/a", "none", "NONE", "nan", "NAN"]: 
            return ""
        return s
    except Exception:
        return ""
