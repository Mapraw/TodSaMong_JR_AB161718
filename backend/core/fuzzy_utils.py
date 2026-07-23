import re
from typing import Dict
from thefuzz import process as fuzz_process

def robust_clean(text: str) -> str:
    """
    Aggressively cleans strings for matching by removing common business suffixes,
    prefixes, and special characters.
    """
    if not text: return ""
    t = str(text).strip()
    if len(t) == 1 and not t.isalnum(): return ""
    if t.upper() in ["NAN", "N/A", "-", ".", ")", "("]: return ""
    
    t = t.upper()
    
    # Remove complex descriptors first
    t = re.sub(r'\b(THE CONSORTIUM OF|CONSORTIUM OF|JOINT VENTURE OF|JOINT VENTURE|CONSORTIUM)\b', '', t)
    
    # Remove common corporate suffixes and geographic noise
    t = re.sub(r'\b(CO|LTD|CORP|INC|LIMITED|COMPANY|PLC|CO\., LTD|PUBLIC|THAILAND|P\.R CHINA|VIETNAM)\b', '', t)
    
    # Remove special characters and replace with space
    t = re.sub(r'[\.\,\/\-\(\)\&]', ' ', t)
    
    # Normalize spacing
    return " ".join(t.split())

def fuzzy_match_name(name: str, choices_dict: Dict, threshold: int = 85) -> str:
    """
    Matches a name against a dictionary of choices using fuzzy string matching.
    """
    clean_name = robust_clean(name)
    if not clean_name: return name
    if clean_name in choices_dict: return choices_dict[clean_name]
    
    choices = list(choices_dict.keys())
    if not choices: return name
    
    match, score = fuzz_process.extractOne(clean_name, choices)
    return choices_dict[match] if score >= threshold else name
