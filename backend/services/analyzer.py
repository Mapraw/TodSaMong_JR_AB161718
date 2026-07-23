from ..core.excel_utils import get_cell_val
from ..config import IDENTITY_CELLS

def identify_category(ws) -> str:
    """
    Analyzes a worksheet to determine if it belongs to AB16, AB17, or AB18.
    Returns the category key (e.g., 'AB18') or None if not recognized.
    """
    # Check for summary or total price keywords to skip these sheets
    a1_val = str(get_cell_val(ws, "A1") or "").upper()
    a2_val = str(get_cell_val(ws, "A2") or "").upper()
    
    # Check AB18 first (most specific header cell)
    ab18_config = IDENTITY_CELLS["AB18"]
    if ab18_config["header_text"] in a2_val:
        return "AB18"
    
    # Identify AB16 first as it might contain AB17 keywords in description
    if IDENTITY_CELLS["AB16_TERM"]["header_text"].upper() in a1_val:
        return "AB16_TERM"
    
    if IDENTITY_CELLS["AB16_CLEAT"]["header_text"].upper() in a1_val:
        return "AB16_CLEAT"

    # Identify AB17
    if IDENTITY_CELLS["AB17"]["header_text"].upper() in a1_val:
        return "AB17"

    return None

def get_schedule_no(ws, category: str) -> str:
    """
    Retrieves the schedule number from the sheet based on its category.
    """
    if not category:
        return ""
        
    config = IDENTITY_CELLS.get(category)
    if config:
        return get_cell_val(ws, config["id_cell"])
        
    return ""
