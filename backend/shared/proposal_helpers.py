import io
import re

import pandas as pd

from ..config import MASTER_COUNTRIES


def normalize_id(text: str) -> str:
    if not text:
        return ""
    return re.sub(r"[^0-9A-Z]+", "", str(text).upper())


def extract_metadata_from_price_schedule(price_bytes):
    items_map = {}
    try:
        xlsx = pd.ExcelFile(io.BytesIO(price_bytes))
        for sheet_name in xlsx.sheet_names:
            if sheet_name.startswith("_") and sheet_name != "_SUM ALL":
                continue
            df = xlsx.parse(sheet_name, header=None)
            for i in range(len(df)):
                row = df.iloc[i]
                for col_idx, cell_val in enumerate(row):
                    val_str = str(cell_val).strip()
                    if re.match(r"\d+AB\d+-\d+", val_str, re.I):
                        item_id = val_str
                        description = ""
                        for offset in range(1, 10):
                            if col_idx + offset < len(row):
                                potential_desc = str(row.iloc[col_idx + offset]).strip()
                                if potential_desc and potential_desc.lower() != "nan" and len(potential_desc) > 10:
                                    description = potential_desc
                                    break
                        clean_key = normalize_id(item_id)
                        if clean_key:
                            items_map[clean_key] = {
                                "schedule": re.search(r"\d+AB\d+", item_id, re.I).group(0).upper(),
                                "item_no": item_id,
                                "description": description,
                            }
        return items_map
    except Exception as exc:
        print(f"Error: {exc}")
        return {}


def extract_item_id_from_pd(df_pd):
    for i in range(min(40, len(df_pd))):
        row_str = " ".join(str(v) for v in df_pd.iloc[i].values if pd.notna(v))
        match = re.search(r"\d+AB\d+-\d+", row_str, re.I)
        if match:
            return match.group(0).upper()
    return None


def extract_detailed_manufacturer_from_pd(df_pd):
    labels_to_strip = [
        r"Manufacturer and Country of Origin",
        r"Manufacturer / Country",
        r"Manufacturer",
        r"Country of origin",
        r"Country",
        r"Origin",
        r"Type / Model / Cat. NO.",
        r"Type / Model",
        r"Type",
        r"Model",
        r"Cat\. NO\.",
        r"Catalogue No\.",
        r"Cable Terminations",
        r"Cable Termination",
        r"Cable Cleats",
        r"Cable Cleat",
    ]
    whitelist = ["Bangkok", "Cable", "Co.,", "Ltd.", "Precise", "CV(TS)", "Thailand"]

    extracted = {}
    markers = {}
    for idx, row in df_pd.iterrows():
        row_str = " ".join(str(v) for v in row.values if pd.notna(v)).strip()
        match = re.match(r"^\s*([a-b])[\.\)]", row_str, re.I)
        if match:
            markers[match.group(1).lower()] = idx

    for marker, start_idx in markers.items():
        combined_text = ""
        for offset in range(4):
            curr_idx = start_idx + offset
            if curr_idx >= len(df_pd):
                break

            row_str = " ".join(str(v) for v in df_pd.iloc[curr_idx].values if pd.notna(v)).strip()
            if offset > 0 and re.match(r"^\s*[a-s][\.\)]", row_str, re.I):
                break

            row_data = df_pd.iloc[curr_idx]
            for value in row_data.values:
                if pd.notna(value):
                    cell_text = str(value).strip()
                    if cell_text:
                        combined_text += " " + cell_text

        cleaned = re.sub(rf"^\s*{marker}[\.\)]", "", combined_text, flags=re.I).strip()
        for pattern in labels_to_strip:
            cleaned = re.sub(pattern, "", cleaned, flags=re.I).strip()

        cleaned = re.sub(r"^[:\-\s/\|]+", "", cleaned).strip()
        cleaned = re.sub(r"[:\-\s/\|]+$", "", cleaned).strip()

        if not cleaned:
            found_whitelisted = []
            for word in whitelist:
                if word.lower() in combined_text.lower():
                    found_whitelisted.append(word)
            if found_whitelisted:
                cleaned = " ".join(dict.fromkeys(found_whitelisted))

        cleaned = re.sub(r"\s+", " ", cleaned)
        if cleaned:
            extracted[marker] = cleaned

    raw_a = extracted.get("a", "")
    raw_b = extracted.get("b", "")

    def parse_manufacturer_country(text: str):
        if not text:
            return "", ""

        text = re.sub(r"\s+", " ", text).strip()

        for delim in [r"\s+-\s+", r"\s+/\s+", r"\s*,\s+"]:
            parts = re.split(delim, text)
            if len(parts) >= 2:
                candidate_country = parts[-1].strip()
                cand_upper = candidate_country.upper()
                common_countries = {
                    "THAILAND",
                    "CHINA",
                    "INDIA",
                    "JAPAN",
                    "KOREA",
                    "VIETNAM",
                    "GERMANY",
                    "USA",
                    "SWEDEN",
                    "SWITZERLAND",
                    "ITALY",
                    "FRANCE",
                    "SPAIN",
                }
                if cand_upper in MASTER_COUNTRIES or cand_upper in common_countries:
                    manufacturer = " ".join(parts[:-1]).strip()
                    country = MASTER_COUNTRIES.get(cand_upper, candidate_country.upper())
                    return manufacturer, country

        common_countries = [
            "THAILAND",
            "CHINA",
            "INDIA",
            "JAPAN",
            "KOREA",
            "VIETNAM",
            "GERMANY",
            "USA",
            "SWEDEN",
            "SWITZERLAND",
            "ITALY",
            "FRANCE",
            "SPAIN",
        ]
        text_upper = text.upper()
        found_country = None
        found_orig_name = None

        all_known = list(MASTER_COUNTRIES.keys()) + list(MASTER_COUNTRIES.values()) + common_countries
        all_known = sorted(set(all_known), key=len, reverse=True)

        for country_name in all_known:
            pattern = rf"\b{re.escape(country_name)}\b"
            if re.search(pattern, text_upper):
                found_country = MASTER_COUNTRIES.get(country_name, country_name).upper()
                found_orig_name = country_name
                break

        if found_country:
            match = re.search(rf"\b{re.escape(found_orig_name)}\b", text_upper)
            start_idx = match.start()
            end_idx = match.end()

            prefix = text[:start_idx].strip()
            prefix = re.sub(r"[:\-\s/\|,\(\)]+$", "", prefix).strip()

            suffix = text[end_idx:].strip()
            suffix = re.sub(r"^[:\-\s/\|,\(\)]+", "", suffix).strip()

            manufacturer = f"{prefix} {suffix}".strip()
            manufacturer = re.sub(r"\s+", " ", manufacturer)
            return manufacturer, found_country

        for delim in [r"\s+-\s+", r"\s+/\s+"]:
            parts = re.split(delim, text)
            if len(parts) >= 2:
                manufacturer = " ".join(parts[:-1]).strip()
                country = parts[-1].strip().upper()
                country = MASTER_COUNTRIES.get(country, country)
                return manufacturer, country

        return text, ""

    mfr_name, country_name = parse_manufacturer_country(raw_a)
    type_model_name = raw_b
    return {
        "manufacturer": mfr_name or "Unknown",
        "country": country_name or "",
        "type_model": type_model_name or "",
    }
