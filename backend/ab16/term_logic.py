import os
import re
import pandas as pd
from ..shared.standard_paths import category_standard_path

def extract_fields_from_pd(df):
    fields = {}
    
    # Initialize with empty strings
    expected_keys = [
        "Manufacturer", "Country of origin", "Type / Model", "Applied Standard", 
        "Voltage Class", "Rated Voltage", "Rated Voltage Note", "BIL", "Power frequency", 
        "Partial discharge", "Operating Environments", "Aerial Connector", "XLPE Suitability",
        "Material of Insulator", "Applied Standard Note",
        "Insulation Diameter Max", "Insulation Diameter Min", "Item Code",
        "f_cond_od", "g_cond_screen", "h_insulation"
    ]
    for key in expected_keys:
        fields[key] = ""

    def get_val_near(r, c, max_offset=30, allow_next_row=True):
        vals = []
        # หาในบรรทัดเดียวกันทางขวา
        for offset in range(1, max_offset):
            if c + offset < len(df.columns):
                v = str(df.iloc[r, c + offset]).strip()
                if v and v.lower() not in {'nan', 'none'} and "PLEASE SPECIFY" not in v.upper():
                    vals.append(v)
        
        # ถ้าหาบรรทัดเดียวกันไม่เจอ ลองหาบรรทัดถัดไปข้างล่าง (ร่นลงมา 1 บรรทัด)
        if allow_next_row and not vals and r + 1 < len(df):
            for offset in range(1, max_offset):
                if c + offset < len(df.columns):
                    v = str(df.iloc[r + 1, c + offset]).strip()
                    if v and v.lower() not in {'nan', 'none'} and "PLEASE SPECIFY" not in v.upper():
                        vals.append(v)
                        
        return " ".join(vals) if vals else ""

    # Helper: Set if empty only
    def set_f(key, val, r, c, lbl):
        if not fields.get(key) and val:
            print(f"  [Parser] Found {lbl} at [{r},{c}], Value: '{val}'")
            fields[key] = val

    for r in range(len(df)):
        for c in range(len(df.columns)):
            val = str(df.iloc[r, c]).strip()
            if not val or val == 'nan': continue
            
            # Allow "IN CASE OF OTHER" to be processed so we can see it
            if "PLEASE SPECIFY" in val.upper() or "BOX BELOW" in val.upper():
                continue

            val_up = val.upper()
            
            if "MANUFACTURER" in val_up and "COUNTRY" not in val_up:
                set_f("Manufacturer", get_val_near(r, c, allow_next_row=False), r, c, "Mfr")
            elif "COUNTRY OF ORIGIN" in val_up:
                set_f("Country of origin", get_val_near(r, c, allow_next_row=False), r, c, "Country")
            elif "TYPE / MODEL" in val_up or "TYPE/MODEL" in val_up:
                type_val = get_val_near(r, c)
                set_f("Type / Model", type_val, r, c, "Model")
                set_f("Item Code", type_val, r, c, "Code")
            elif "OUTSIDE DIAMETER OF STRANDED CONDUCTOR" in val_up:
                set_f("f_cond_od", get_val_near(r, c), r, c, "f-Cond")
            elif "CONDUCTOR SCREEN" in val_up:
                set_f("g_cond_screen", get_val_near(r, c), r, c, "g-Screen")
            elif "INSULATION" in val_up and "PARTIAL" not in val_up and "DIAMETER" not in val_up:
                # 1AB17 Format: "Insulation" is a main header, and the next row might have "(Material / Thickness)"
                if r + 1 < len(df):
                    next_row_val = str(df.iloc[r+1, c]).strip().upper()
                    # It might be in the same column or slightly offset
                    if not next_row_val or next_row_val == 'NAN':
                         for offset in range(1, 5):
                             if c + offset < len(df.columns):
                                 check_val = str(df.iloc[r+1, c+offset]).strip().upper()
                                 if "THICKNESS" in check_val:
                                     next_row_val = check_val
                                     break
                    
                    if "THICKNESS" in next_row_val:
                         # Found the thickness row, extract from there
                         set_f("h_insulation", get_val_near(r+1, c), r+1, c, "h-Ins")
                    else:
                         # Fallback to current row if format differs
                         set_f("h_insulation", get_val_near(r, c), r, c, "h-Ins")
                else:
                    set_f("h_insulation", get_val_near(r, c), r, c, "h-Ins")
            elif "APPLIED STANDARD" in val_up:
                set_f("Applied Standard", get_val_near(r, c), r, c, "Std")
                if r + 1 < len(df):
                    # Capture the note field content specifically
                    # We check the cell at (r+1, c) as well
                    note_label = str(df.iloc[r+1, c]).strip()
                    note_val = get_val_near(r + 1, c)
                    print(f"  [Parser Debug] Applied Std Row {r}, Note Row {r+1}, Label: '{note_label}', Value: '{note_val}'")
                    if note_val:
                        fields["Applied Standard Note"] = note_val
                    elif "IN CASE OF OTHER" in note_label.upper():
                        # If the value is empty but the label matches, we still might want to track it
                        # But user wants to warn if DETECTED in the box
                        pass
            elif "VOLTAGE CLASS" in val_up or "SYSTEM VOLTAGE" in val_up or "NOMINAL VOLTAGE" in val_up:
                set_f("Voltage Class", get_val_near(r, c), r, c, "V-Class")
            elif "RATED VOLTAGE" in val_up:
                near_val = get_val_near(r, c)
                set_f("Rated Voltage", near_val, r, c, "Rated-V")
                set_f("Voltage Class", near_val, r, c, "V-Class-Fallback") # ใช้ Rated V แทน Voltage Class
                if r + 1 < len(df):
                    note_val = get_val_near(r + 1, c)
                    if note_val:
                        fields["Rated Voltage Note"] = note_val
            elif "IMPULSE WITHSTAND VOLTAGE" in val_up or re.search(r"\bBIL\b", val_up):
                near_val = get_val_near(r, c)
                # ดักจับกรณีที่ผู้ใช้ลบตัวเลขไปแล้ว แต่เหลือคำว่า "kV" ทิ้งไว้ในช่องถัดไป
                if near_val.lower() == "kv" or near_val.strip() == "":
                    near_val = ""
                set_f("BIL", near_val, r, c, "BIL")
            elif "POWER FREQUENCY" in val_up:
                set_f("Power frequency", get_val_near(r, c), r, c, "P-Freq")
            elif "PARTIAL DISCHARGE" in val_up:
                set_f("Partial discharge", get_val_near(r, c), r, c, "PD")
            elif "OPERATING ENVIRONMENTS" in val_up:
                set_f("Operating Environments", get_val_near(r, c), r, c, "Environ")
            elif "AERIAL CONNECTOR" in val_up:
                set_f("Aerial Connector", get_val_near(r, c), r, c, "Connector")
            elif "XLPE" in val_up:
                set_f("XLPE Suitability", get_val_near(r, c), r, c, "XLPE")
            elif "MATERIAL" in val_up and "INSULATOR" in val_up:
                set_f("Material of Insulator", get_val_near(r, c), r, c, "Material")
            elif "INSULATION" in val_up and "DIAMETER" in val_up:
                near_val = get_val_near(r, c)
                nums = re.findall(r"(\d+(?:\.\d+)?)", near_val)
                # ถ้าเจอตัวเลข 2 ตัว ให้จับแยก Min/Max เลย ไม่ต้องสนใจคำว่า MIN/MAX ในหัวข้อ
                if len(nums) >= 2:
                    n1, n2 = float(nums[0]), float(nums[1])
                    set_f("Insulation Diameter Min", str(min(n1, n2)), r, c, "Ins-Min-Auto")
                    set_f("Insulation Diameter Max", str(max(n1, n2)), r, c, "Ins-Max-Auto")
                else:
                    if "MAX" in val_up:
                        set_f("Insulation Diameter Max", near_val, r, c, "Ins-Max")
                    elif "MIN" in val_up:
                        set_f("Insulation Diameter Min", near_val, r, c, "Ins-Min")
                    elif not fields["Insulation Diameter Max"]:
                        set_f("Insulation Diameter Max", near_val, r, c, "Ins-Near")
                set_f("XLPE Suitability", get_val_near(r, c), r, c, "XLPE")
            elif "MATERIAL" in val_up and "INSULATOR" in val_up:
                set_f("Material of Insulator", get_val_near(r, c), r, c, "Material")
                
    return fields

_STANDARD_CACHE = {}

def _clean_text(value):
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    return text

def _norm_text(value):
    return re.sub(r"[^A-Z0-9]+", "", _clean_text(value).upper())

def _numbers_from_text(value):
    return [float(n) for n in re.findall(r"(\d+(?:\.\d+)?)", _clean_text(value))]

def _explicit_kv_numbers(value):
    return [float(n) for n in re.findall(r"(\d+(?:\.\d+)?)\s*kV", _clean_text(value), re.I)]

def _measurement_numbers(value):
    numbers = _explicit_kv_numbers(value)
    if numbers:
        return numbers
    return _numbers_from_text(value)

def _close_to(value, target):
    return abs(float(value) - float(target)) < 0.01

def _standard_path(file_name):
    return category_standard_path("ab16", file_name)

def _dielectric_standard_path():
    return _standard_path("g. & h. standard files.xlsx")

def _rf_standard_path():
    return _standard_path("i. & j. .xlsx")

def _map_to_system_voltage(number):
    for system_kv in [11, 22, 33, 69, 115, 230, 500]:
        if _close_to(number, system_kv):
            return system_kv
    if 20 <= number <= 25:
        return 22
    if 30 <= number <= 36:
        return 33
    if 65 <= number <= 72:
        return 69
    if 110 <= number <= 120:
        return 115
    if 220 <= number <= 245:
        return 230
    if 480 <= number <= 525:
        return 500
    return None

def _infer_system_voltage_kv(pd_fields, ps_description):
    for source in [ps_description, pd_fields.get("Voltage Class", "")]:
        for number in _explicit_kv_numbers(source):
            system_kv = _map_to_system_voltage(number)
            if system_kv:
                return system_kv

    for number in _numbers_from_text(pd_fields.get("Voltage Class", "")):
        system_kv = _map_to_system_voltage(number)
        if system_kv:
            return system_kv

    for number in _explicit_kv_numbers(pd_fields.get("Rated Voltage", "")):
        if _close_to(number, 25):
            return 22
        if _close_to(number, 35):
            return 33
        system_kv = _map_to_system_voltage(number)
        if system_kv:
            return system_kv
    return None

def _infer_insulation_class_kv(pd_fields, ps_description, system_kv=None):
    valid_classes = [2.5, 5, 8, 15, 25, 35, 46]
    rated_voltage = pd_fields.get("Rated Voltage", "")
    sources = []
    if "OTHER" in _norm_text(rated_voltage):
        sources.append(pd_fields.get("Rated Voltage Note", ""))
    sources.extend([rated_voltage, pd_fields.get("Voltage Class", ""), ps_description])
    for source in sources:
        for number in _explicit_kv_numbers(source):
            for class_kv in valid_classes:
                if _close_to(number, class_kv):
                    return class_kv
            nearest = min(valid_classes, key=lambda class_kv: abs(class_kv - number))
            if abs(nearest - number) <= 3:
                return nearest

    if system_kv == 22:
        return 25
    if system_kv == 33:
        return 35
    return None

def _rated_voltage_class_kv(rated_voltage):
    valid_classes = [2.5, 5, 8, 15, 25, 35, 46]
    for number in _measurement_numbers(rated_voltage):
        for class_kv in valid_classes:
            if _close_to(number, class_kv):
                return class_kv
    return None

def _load_bil_standard():
    if "bil" in _STANDARD_CACHE:
        return _STANDARD_CACHE["bil"]

    standard = {}
    path = _standard_path("f. (BIL).xlsx")
    if os.path.exists(path):
        df = pd.read_excel(path, header=None)
        for _, row in df.iterrows():
            if len(row) < 2:
                continue
            bil_nums = _numbers_from_text(row.iloc[0])
            system_nums = _numbers_from_text(row.iloc[1])
            if bil_nums and system_nums:
                standard[int(system_nums[0])] = bil_nums[0]

    _STANDARD_CACHE["bil"] = standard
    return standard

def _load_dielectric_standard():
    if "dielectric" in _STANDARD_CACHE:
        return _STANDARD_CACHE["dielectric"]

    standard = {}
    preferred_path = _dielectric_standard_path()
    local_path = _standard_path("g. & h. standard files.xlsx")
    paths = [preferred_path]
    if os.path.normcase(preferred_path) != os.path.normcase(local_path):
        paths.append(local_path)

    for path in paths:
        if not os.path.exists(path):
            continue
        try:
            df = pd.read_excel(path, header=None)
        except OSError:
            continue
        for _, row in df.iterrows():
            if len(row) < 3:
                continue
            class_nums = _numbers_from_text(row.iloc[0])
            if len(row) >= 4:
                pd_nums = _numbers_from_text(row.iloc[2])
                dry_nums = _numbers_from_text(row.iloc[3])
            else:
                pd_nums = _numbers_from_text(row.iloc[1])
                dry_nums = _numbers_from_text(row.iloc[2])
            if class_nums and pd_nums and dry_nums:
                standard[class_nums[0]] = {"pd": pd_nums[0], "dry": dry_nums[0]}
        if standard:
            break

    _STANDARD_CACHE["dielectric"] = standard
    return standard

def _load_rf_standard():
    if "rf" in _STANDARD_CACHE:
        return _STANDARD_CACHE["rf"]

    standard = {}
    preferred_path = _rf_standard_path()
    local_path = _standard_path("i. & j. .xlsx")
    paths = [preferred_path]
    if os.path.normcase(preferred_path) != os.path.normcase(local_path):
        paths.append(local_path)

    for path in paths:
        if not os.path.exists(path):
            continue
        try:
            df = pd.read_excel(path, header=None)
        except OSError:
            continue
        for _, row in df.iterrows():
            values = [_clean_text(v) for v in row.tolist()]
            values = [v for v in values if v]
            if len(values) < 3:
                continue
            code_match = re.match(r"([A-Z]{2}\d+[A-Z0-9]*)", values[0].upper())
            if code_match:
                standard[code_match.group(1)] = {
                    "environment": values[1],
                    "connector": values[2],
                }
        if standard:
            break

    _STANDARD_CACHE["rf"] = standard
    return standard

def _parse_k_range(value):
    match = re.search(
        r"(\d+(?:\.\d+)?)\s*[-–—]\s*(\d+(?:\.\d+)?)",
        _clean_text(value),
    )
    if not match:
        return None
    range_min, range_max = float(match.group(1)), float(match.group(2))
    return min(range_min, range_max), max(range_min, range_max)

def _k_model_aliases(value):
    text = _clean_text(value).upper()
    aliases = set()
    full_norm = _norm_text(text)
    if full_norm:
        aliases.add(full_norm)

    for match in re.findall(r"TFT\s*-?\s*\d+[A-Z]?(?:\s*-\s*T\d+)?", text, re.I):
        aliases.add(_norm_text(match))

    for match in re.findall(r"\b\d{4}(?:\s*-\s*[A-Z0-9]+){0,2}\b", text, re.I):
        aliases.add(_norm_text(match))

    return aliases

def _k_voltage_classes(value):
    classes = set()
    for match in re.finditer(
        r"(\d+(?:\.\d+)?)(?:\s*-\s*(\d+(?:\.\d+)?))?\s*kV\s*Class",
        _clean_text(value),
        re.I,
    ):
        classes.add(float(match.group(1)))
        if match.group(2):
            classes.add(float(match.group(2)))
    return classes

def _k_standard_filename(type_model):
    model_norm = _norm_text(type_model)
    if "TFT" in model_norm:
        return "k. TFT.xlsx"
    if "QTIII" in model_norm:
        return "k. 3M QT-III.xlsx"
    if "QTII" in model_norm:
        return "k. 3M QT-II.xlsx"
    return ""

def _load_k_standard_entries(file_name):
    cache_key = f"k:{file_name}"
    if cache_key in _STANDARD_CACHE:
        return _STANDARD_CACHE[cache_key]

    entries = []
    path = _standard_path(file_name)
    if not os.path.exists(path):
        _STANDARD_CACHE[cache_key] = entries
        return entries

    df = pd.read_excel(path, header=None)
    range_columns = set()
    model_columns = set()
    voltage_column = None

    for _, row in df.iterrows():
        for column_index, value in enumerate(row.tolist()):
            value_norm = _norm_text(value)
            if "INSULATIONOD" in value_norm and ("RANGE" in value_norm or "MINMAX" in value_norm):
                range_columns.add(column_index)
            if value_norm == "DESCRIPTION" or value_norm.startswith("INDOOR") or value_norm.startswith("OUTDOOR"):
                model_columns.add(column_index)
            if value_norm == "VOLTAGEKV":
                voltage_column = column_index

    current_classes = set()
    for _, row in df.iterrows():
        if voltage_column is not None and voltage_column < len(row):
            row_classes = _k_voltage_classes(row.iloc[voltage_column])
            if row_classes:
                current_classes = row_classes

        od_range = None
        for range_column in range_columns:
            if range_column < len(row):
                od_range = _parse_k_range(row.iloc[range_column])
                if od_range:
                    break
        if not od_range:
            continue

        aliases = set()
        display_models = []
        for model_column in model_columns:
            if model_column >= len(row):
                continue
            model_value = _clean_text(row.iloc[model_column])
            if not model_value:
                continue
            model_aliases = _k_model_aliases(model_value)
            if model_aliases:
                aliases.update(model_aliases)
                display_models.append(model_value)

        if aliases:
            entries.append({
                "aliases": aliases,
                "models": display_models,
                "min": od_range[0],
                "max": od_range[1],
                "classes": set(current_classes),
            })

    _STANDARD_CACHE[cache_key] = entries
    return entries

def _lookup_k_standard(type_model, rated_voltage):
    file_name = _k_standard_filename(type_model)
    if not file_name:
        return "", None

    model_aliases = _k_model_aliases(type_model)
    candidates = [
        entry for entry in _load_k_standard_entries(file_name)
        if model_aliases.intersection(entry["aliases"])
    ]
    if not candidates:
        return file_name, None

    rated_class = _rated_voltage_class_kv(rated_voltage)
    if rated_class is not None:
        class_matches = [
            entry for entry in candidates
            if not entry["classes"] or rated_class in entry["classes"]
        ]
        if class_matches:
            candidates = class_matches

    return file_name, candidates[0]

def _extract_rf_code(ps_description):
    standard = _load_rf_standard()
    description = _clean_text(ps_description)
    description_norm = _norm_text(ps_description)

    rf_match = re.search(
        r"\bRF(?:\s*(?:NO\.?|CODE))?\s*[-:]?\s*([A-Z0-9]+)\b",
        description,
        re.I,
    )
    if rf_match:
        rf_token = _norm_text(rf_match.group(1))
        for code in standard:
            code_norm = _norm_text(code)
            if rf_token == code_norm:
                return code
            if code_norm.startswith("TN") and code_norm.endswith("H"):
                rf_core = code_norm[2:-1]
                if rf_token in {rf_core, f"TN{rf_core}", f"{rf_core}H"}:
                    return code

    for code in sorted(standard, key=len, reverse=True):
        if code in description_norm:
            return code
    return ""

def _matches_standard_text(actual, expected):
    actual_text = re.sub(r"\s+", " ", _clean_text(actual)).casefold()
    expected_text = re.sub(r"\s+", " ", _clean_text(expected)).casefold()
    return bool(actual_text) and actual_text == expected_text

def _has_expected_kv(value, expected):
    return any(_close_to(number, expected) for number in _measurement_numbers(value))

def _has_at_least_kv(value, expected):
    numbers = _measurement_numbers(value)
    return bool(numbers) and max(numbers) >= expected

def _has_one_minute(value):
    text = _clean_text(value)
    return bool(re.search(r"\b1(?:\.0+)?\s*(?:min|minute)\b", text, re.I) or re.search(r"\b1(?:\.0+)?\s*นาที\b", text))

def _extract_power_frequency_parts(value):
    text = _clean_text(value)
    kv_match = re.search(r"(\d+(?:\.\d+)?)\s*kV", text, re.I)
    minute_match = re.search(r"(\d+(?:\.\d+)?)\s*(?:min|minute|นาที)\b", text, re.I)

    numbers = _numbers_from_text(text)
    kv_value = float(kv_match.group(1)) if kv_match else (numbers[0] if numbers else None)
    minute_value = float(minute_match.group(1)) if minute_match else (numbers[1] if len(numbers) > 1 else None)
    return kv_value, minute_value

def _meaningful_standard_note(value):
    text = _clean_text(value)
    if not text:
        return ""
    note_up = text.upper()
    ignored = ["PLEASE SPECIFY", "IN CASE OF OTHER", "BOX BELOW"]
    if any(keyword in note_up for keyword in ignored):
        return ""
    return text

def _keywords_match(actual, expected):
    actual_norm = _norm_text(actual)
    words = re.findall(r"[A-Z0-9]+", _clean_text(expected).upper())
    words = [word for word in words if len(word) > 1 and word not in {"REV", "CONFORM", "TO"}]
    return bool(words) and all(word in actual_norm for word in words)

def _connector_option(value):
    norm = _norm_text(value)
    if norm in {"OTHER", "OTHERS"}:
        return "OTHER"
    if any(token in norm for token in ["2X3", "6HOLE", "6BOLT"]):
        return "6"
    if "CLAMP" in norm:
        return "CLAMP"
    if any(token in norm for token in ["2X2", "4HOLE", "4BOLT"]):
        return "4"
    if any(token in norm for token in ["2X1", "2HOLE", "2BOLT"]):
        return "2"
    return ""

def _connector_matches_standard(actual, expected):
    actual_option = _connector_option(actual)
    expected_option = _connector_option(expected)
    if actual_option in {"2", "4"}:
        return actual_option == expected_option
    return _matches_standard_text(actual, expected)

def _is_supported_nema_connector(value):
    option = _connector_option(value)
    norm = _norm_text(value)
    return option in {"2", "4"} and ("NEMA" in norm or "TERMINALPAD" in norm or "PAD" in norm)

def check_qwerty_logic(pd_fields, ps_description, ab17_min=None, ab17_max=None, ref_fields=None):
    error_topics = []
    error_details = []      # บอกว่า "ผิดยังไง"
    fix_suggestions = []    # บอกว่า "ต้องแก้ยังไง"
    
    print(f"  [Checking Logic] PS Description: '{ps_description}'")
    if ab17_min is not None and ab17_max is not None:
        print(f"  [Checking Logic] Reference Value from 1AB17: {ab17_min} - {ab17_max}")

    # ดึงค่า f, g, h จากไฟล์อ้างอิง (1AB17) มาเติมให้ 1AB16 แบบอัตโนมัติ เพื่อกันไม่ให้ฟ้องว่าว่าง
    if ref_fields:
        if not pd_fields.get("f_cond_od"):
            pd_fields["f_cond_od"] = ref_fields.get("f_cond_od", "")
        if not pd_fields.get("g_cond_screen"):
            pd_fields["g_cond_screen"] = ref_fields.get("g_cond_screen", "")
        if not pd_fields.get("h_insulation"):
            pd_fields["h_insulation"] = ref_fields.get("h_insulation", "")

    # ตรวจสอบข้อมูลว่างเปล่า (Empty Fields Check)
    missing_fields = []
    
    # รายชื่อฟิลด์ที่ผูกกับเงื่อนไขข้อ a-k (ระบุตัวอักษรให้ตรงกับกฎ)
    required_fields_map = {
        "Manufacturer": "a. Manufacturer",
        "Country of origin": "a. Country of origin",
        "Type / Model": "- Type / Model (รุ่นสินค้า)",
        "Applied Standard": "c. Applied Standard",
        "Material of Insulator": "d. Material of Insulator",
        "Rated Voltage": "e. Rated Voltage",
        "BIL": "f. Impulse Withstand Voltage (BIL)",
        "Power frequency": "g. Power frequency (Dry withstand)",
        "Partial discharge": "h. Partial discharge",
        "Operating Environments": "i. Operating Environments (Creepage)",
        "Aerial Connector": "j. Aerial Connector",
        "Insulation Diameter Min": "k. Insulation Diameter (Min)",
        "Insulation Diameter Max": "k. Insulation Diameter (Max)"
    }

    # สำหรับไฟล์สายเคเบิล (1AB17) ถึงจะบังคับตรวจข้อ f, g, h ของสายเคเบิล
    # ถ้าเป็น 1AB16 (หัวเคเบิล) จะไม่บังคับตรวจ 3 ข้อนี้ เพราะไม่มีให้กรอก
    cable_fields = {
        "f_cond_od": "f. Outside Diameter of Stranded Conductor (สายเคเบิล)",
        "g_cond_screen": "g. Conductor Screen (สายเคเบิล)",
        "h_insulation": "h. Insulation Thickness (สายเคเบิล)"
    }
    
    if "1AB17" in ps_description.upper() or "CABLE" in ps_description.upper() and "TERMINATION" not in ps_description.upper():
        required_fields_map.update(cable_fields)

    for dict_key, display_name in required_fields_map.items():
        val = pd_fields.get(dict_key, "").strip()
        
        # ถ้าระบบไปจับได้แค่หน่วย (เช่น "kV", "kV min") แต่ไม่มีตัวเลขเลย ให้ถือว่า "ว่าง" ทันที
        numeric_fields = ["Rated Voltage", "BIL", "Partial discharge", "Insulation Diameter Min", "Insulation Diameter Max"]
        if dict_key in numeric_fields:
            is_rated_voltage_other = dict_key == "Rated Voltage" and "OTHER" in _norm_text(val)
            if not re.search(r'\d', val) and not is_rated_voltage_other:
                val = ""
        
        # กฎพิเศษสำหรับข้อ g (Power frequency) แยกเช็ค kV และ min ให้แม่นยำขึ้น
        if dict_key == "Power frequency":
            has_kv_val = bool(re.search(r'(\d+(?:\.\d+)?)\s*kV', val, re.I))
            has_min_val = bool(re.search(r'(\d+(?:\.\d+)?)\s*(?:min|minute)', val, re.I))
            voltage_display = "g. ยังไม่ได้กรอกค่าแรงดัน Power Frequency (kV)"
            time_display = "g. ยังไม่ได้กรอกเวลาทดสอบ Power Frequency (min/minute)"
            
            if not val or val.lower() == 'nan' or (not has_kv_val and not has_min_val):
                # ว่างหมดเลย (หรือมีแต่ตัวหนังสือ kV min แต่ไม่มีเลข)
                missing_fields.append((voltage_display, voltage_display))
                missing_fields.append((time_display, time_display))
                val = "PARTIAL_EMPTY"
            elif has_kv_val and not has_min_val:
                # มีตัวเลขหน้า kV แต่ไม่มีตัวเลขหน้า min
                missing_fields.append((time_display, time_display))
                val = "PARTIAL_EMPTY" # Flag to prevent standard empty check
            elif not has_kv_val and has_min_val:
                # มีตัวเลขหน้า min แต่ไม่มีตัวเลขหน้า kV
                missing_fields.append((voltage_display, voltage_display))
                val = "PARTIAL_EMPTY"
                
        if val != "PARTIAL_EMPTY" and (not val or val.lower() == 'nan'):
            missing_fields.append((display_name, display_name))

            
    missing_status = []
    missing_comment = []
    
    if missing_fields:
        error_topics.append("missing_data")
        missing_status.append("⚠️ ตรวจพบช่องข้อมูลที่ยังไม่ได้กรอก:")
        for short_name, detailed_name in missing_fields:
            missing_status.append(f"  {short_name}")
            
        missing_comment.append("⚠️ กรุณาเติมข้อมูลในช่องต่อไปนี้ให้ครบถ้วน:")
        for short_name, detailed_name in missing_fields:
            missing_comment.append(f"  {detailed_name}")

    # a. Manufacturer and Country of Origin
    mfr = pd_fields.get("Manufacturer", "").strip()
    country = pd_fields.get("Country of origin", "").strip()
    if mfr and country:
        if "te" in mfr.lower() and "usa" in country.lower():
            error_topics.append("a")
            error_details.append("a. พบผู้ผลิต TE (USA) ซึ่งย้ายฐานการผลิตแล้ว")
            fix_suggestions.append("a. กรุณาเปลี่ยนเป็นฐานการผลิตปัจจุบัน (Mexico) และแนบเอกสาร Type Test ใหม่")

    system_kv = _infer_system_voltage_kv(pd_fields, ps_description)
    insulation_class_kv = _infer_insulation_class_kv(pd_fields, ps_description, system_kv)

    # c. Applied Standard
    applied_standard = pd_fields.get("Applied Standard", "").strip()
    applied_standard_note = _meaningful_standard_note(pd_fields.get("Applied Standard Note", ""))
    applied_standard_norm = _norm_text(applied_standard)
    if applied_standard:
        if "OTHER" in applied_standard_norm:
            if not applied_standard_note:
                error_topics.append("c")
                error_details.append("c. เลือก Applied Standard เป็น 'Other' แต่ยังไม่ได้ระบุ Note")
                fix_suggestions.append("c. กรุณาระบุชื่อมาตรฐานในช่อง Applied Standard Note ให้ชัดเจน")
        elif applied_standard_note:
            error_topics.append("c")
            error_details.append("c. ตรวจพบข้อมูลในช่อง Applied Standard Note โดยที่มีมาตรฐานอยู่แล้ว")
            fix_suggestions.append("c. กรุณาลบข้อมูลในช่อง Applied Standard Note ออกให้ว่างเปล่า หากไม่ได้เลือกมาตรฐานเป็น Other")

    # d. Material of Insulator
    material = pd_fields.get("Material of Insulator", "").strip()
    material_norm = _norm_text(material)
    if material and system_kv:
        if system_kv <= 33:
            if "SILICONERUBBER" not in material_norm:
                error_topics.append("d")
                error_details.append(f"d. วัสดุฉนวนไม่ถูกต้องสำหรับระบบ {system_kv}kV")
                fix_suggestions.append("d. กรุณาระบุ Material of Insulator เป็น Silicone rubber")
        elif system_kv in {230, 500}:
            valid_material = (
                "HIGHSTRENGTHPORCELAIN" in material_norm
                or "WETPROCESSPORCELAIN" in material_norm
                or "COMPOSITEMATERIAL" in material_norm
            )
            if not valid_material:
                error_topics.append("d")
                error_details.append(f"d. วัสดุฉนวนไม่ถูกต้องสำหรับระบบ {system_kv}kV")
                fix_suggestions.append("d. กรุณาระบุ Material of Insulator เป็น High strength porcelain, Wet process porcelain, หรือ Composite material")

    # e. Rated Voltage
    rated_voltage = pd_fields.get("Rated Voltage", "").strip()
    rated_voltage_note = _meaningful_standard_note(pd_fields.get("Rated Voltage Note", ""))
    rated_voltage_is_other = "OTHER" in _norm_text(rated_voltage)
    rated_numbers = _measurement_numbers(rated_voltage_note if rated_voltage_is_other else rated_voltage)
    if rated_voltage_is_other and not rated_numbers:
        error_topics.append("e")
        error_details.append("e. เลือก Rated Voltage เป็น 'Other' แต่ยังไม่ได้ระบุค่าแรงดัน")
        fix_suggestions.append("e. กรุณาระบุค่าแรงดันในช่องใต้ Other")
    elif rated_voltage and system_kv:
        if not rated_numbers or not any(number >= system_kv for number in rated_numbers):
            error_topics.append("e")
            error_details.append(f"e. Rated Voltage ต่ำกว่าแรงดันระบบ {system_kv}kV")
            fix_suggestions.append(f"e. กรุณาระบุ Rated Voltage ให้มากกว่าหรือเท่ากับ {system_kv}kV")

    # f. BIL (Impulse Withstand Voltage)
    bil = pd_fields.get("BIL", "").strip()
    bil_standard = _load_bil_standard()
    expected_bil = bil_standard.get(system_kv)
    if bil and system_kv and expected_bil:
        if not _has_expected_kv(bil, expected_bil):
            error_topics.append("f")
            error_details.append(f"f. ค่า BIL ไม่ตรงตามมาตรฐานสำหรับระบบ {system_kv}kV")
            fix_suggestions.append(f"f. แก้ไขค่า Impulse Withstand Voltage (BIL) ให้เป็น {expected_bil:g}kV")

    # g. Power Frequency (Dry Withstand)
    power_frequency = pd_fields.get("Power frequency", "").strip()
    dielectric_standard = _load_dielectric_standard()
    dielectric_ref = dielectric_standard.get(insulation_class_kv) if insulation_class_kv else None
    if power_frequency and dielectric_ref:
        expected_dry = dielectric_ref["dry"]
        expected_minutes = 1.0
        power_kv_match = re.search(r'(\d+(?:\.\d+)?)\s*kV', power_frequency, re.I)
        power_minutes_match = re.search(r'(\d+(?:\.\d+)?)\s*(?:min|minute)', power_frequency, re.I)

        if power_kv_match:
            actual_kv = float(power_kv_match.group(1))
            if not _close_to(actual_kv, expected_dry):
                error_topics.append("g")
                error_details.append(f"g. กรอกค่าแรงดัน Power Frequency ไม่ถูกต้อง (ต้องเป็น {expected_dry:g}kV)")
                fix_suggestions.append(f"g. แก้ค่าแรงดันเป็น {expected_dry:g}kV")

        if power_minutes_match:
            actual_minutes = float(power_minutes_match.group(1))
            if not _close_to(actual_minutes, expected_minutes):
                error_topics.append("g")
                error_details.append("g. กรอกเวลาทดสอบ Power Frequency ไม่ถูกต้อง (ต้องเป็น 1 minute)")
                fix_suggestions.append("g. แก้เวลาทดสอบเป็น '1 minute'")

    # h. Partial Discharge
    partial_discharge = pd_fields.get("Partial discharge", "").strip()
    h_insulation_class_kv = _rated_voltage_class_kv(rated_voltage)
    h_dielectric_ref = dielectric_standard.get(h_insulation_class_kv) if h_insulation_class_kv else None
    if partial_discharge and h_dielectric_ref and _measurement_numbers(partial_discharge):
        expected_pd = h_dielectric_ref["pd"]
        if not _has_at_least_kv(partial_discharge, expected_pd):
            error_topics.append("h")
            error_details.append("h. ค่า Partial Discharge ต่ำกว่าเกณฑ์มาตรฐาน")
            fix_suggestions.append(f"h. ปรับค่าให้ไม่น้อยกว่า {expected_pd:g}kV สำหรับ Rated Voltage {h_insulation_class_kv:g}kV")

    # i. Operating Environments (Creepage)
    rf_code = _extract_rf_code(ps_description)
    rf_standard = _load_rf_standard().get(rf_code)
    operating_environment = pd_fields.get("Operating Environments", "").strip()
    if operating_environment:
        if not rf_standard:
            error_topics.append("i")
            error_details.append("i. ไม่พบรหัส RF ที่ตรงกับตารางมาตรฐานจาก Description")
            fix_suggestions.append("i. ตรวจสอบให้ Description ระบุรหัส RF ที่มีอยู่ในไฟล์ i. & j. .xlsx")
        else:
            expected_environment = rf_standard["environment"]
            if not _matches_standard_text(operating_environment, expected_environment):
                error_topics.append("i")
                error_details.append(f"i. Operating Environments ไม่ตรงตาม RF {rf_code}")
                fix_suggestions.append(f"i. แก้ไขค่า Operating Environments ให้เป็น '{expected_environment}' เท่านั้น")

    # j. Aerial Connector
    aerial_connector = pd_fields.get("Aerial Connector", "").strip()
    if aerial_connector:
        if not rf_standard:
            error_topics.append("j")
            error_details.append("j. ไม่พบรหัส RF ที่ตรงกับตารางมาตรฐานจาก Description")
            fix_suggestions.append("j. ตรวจสอบให้ Description ระบุรหัส RF ที่มีอยู่ในไฟล์ i. & j. .xlsx")
        else:
            expected_connector = rf_standard["connector"]
            connector_option = _connector_option(aerial_connector)
            if connector_option == "OTHER":
                error_topics.append("j")
                error_details.append("j. Aerial Connector ระบุเป็น Other แต่ยังไม่ได้ใส่ค่าที่ต้องการ")
                fix_suggestions.append(f"j. กรุณาระบุค่า Aerial Connector ตาม RF {rf_code}: '{expected_connector}'")
            elif connector_option == "6":
                error_topics.append("j")
                error_details.append("j. 2 x 3 bolt holes NEMA Pad ยังไม่มีในตารางมาตรฐาน")
                fix_suggestions.append(f"j. กรุณาตรวจสอบและใช้ค่า Aerial Connector ตาม RF {rf_code}: '{expected_connector}'")
            elif not _connector_matches_standard(aerial_connector, expected_connector):
                error_topics.append("j")
                error_details.append(f"j. Aerial Connector ไม่ตรงตามที่ RF {rf_code} กำหนด")
                fix_suggestions.append(f"j. แก้ไขค่า Aerial Connector ให้เป็น '{expected_connector}' เท่านั้น")

    # k. Insulation Diameter
    type_model = pd_fields.get("Type / Model", "").strip()
    min_numbers = _numbers_from_text(pd_fields.get("Insulation Diameter Min", ""))
    max_numbers = _numbers_from_text(pd_fields.get("Insulation Diameter Max", ""))
    p_min = min_numbers[0] if min_numbers else None
    p_max = max_numbers[0] if max_numbers else None
    standard_file = _k_standard_filename(type_model)

    if type_model and standard_file and p_min is not None and p_max is not None:
        _, k_standard = _lookup_k_standard(type_model, rated_voltage)
        if not k_standard:
            error_topics.append("k")
            error_details.append(f"k. ไม่พบรหัสรุ่นจากข้อ b '{type_model}' ในไฟล์ {standard_file}")
            fix_suggestions.append("k. กรุณาตรวจสอบรหัส Type / Model / Cat. NO. ในข้อ b ให้ตรงกับตารางมาตรฐาน")
        else:
            standard_min = k_standard["min"]
            standard_max = k_standard["max"]
            type_aliases = _k_model_aliases(type_model)
            matched_model = next(
                (
                    model for model in k_standard["models"]
                    if type_aliases.intersection(_k_model_aliases(model))
                ),
                type_model,
            )

            if not _close_to(p_min, standard_min) or not _close_to(p_max, standard_max):
                error_topics.append("k")
                error_details.append(
                    f"k. Insulation Diameter ({p_min:g}-{p_max:g} mm) ไม่ตรงกับรุ่น {matched_model}"
                )
                fix_suggestions.append(
                    f"k. แก้ค่า Min/Max เป็น {standard_min:g}-{standard_max:g} mm ตามไฟล์ {standard_file}"
                )

            if ab17_min is not None and ab17_max is not None:
                if not (p_min <= ab17_min and ab17_max <= p_max):
                    error_topics.append("k")
                    error_details.append(
                        f"k. สายไฟ 1AB17 ({ab17_min:g}-{ab17_max:g} mm) อยู่นอกช่วงของรุ่น {matched_model}"
                    )
                    fix_suggestions.append(
                        f"k. เปลี่ยนรุ่นหัวเคเบิลให้ครอบคลุมขนาด OD {ab17_min:g}-{ab17_max:g} mm"
                    )
            elif ref_fields:
                f_val = ref_fields.get("f_cond_od")
                g_val = ref_fields.get("g_cond_screen")
                h_val = ref_fields.get("h_insulation")
                if f_val and g_val and h_val:
                    try:
                        f_num = _numbers_from_text(f_val)[0]
                        g_num = _numbers_from_text(g_val)[0]
                        h_num = _numbers_from_text(h_val)[0]
                        calc_val = f_num + 2*(g_num + h_num)
                        if not (p_min <= calc_val <= p_max):
                            error_topics.append("k")
                            error_details.append(
                                f"k. สายไฟ 1AB17 (คำนวณได้ {calc_val:.2f} mm) อยู่นอกช่วงของรุ่น {matched_model}"
                            )
                            fix_suggestions.append(
                                f"k. เปลี่ยนรุ่นหัวเคเบิลให้ครอบคลุมขนาด OD {calc_val:.2f} mm"
                            )
                    except Exception:
                        pass

    # --- การรวมข้อความเพื่อแสดงผล ---
    status_parts = []
    comment_parts = []

    # 1. ส่วนของ Logic Errors (ผิดเงื่อนไข a-k)
    if error_details:
        status_parts.append("❌ ไม่ผ่านเงื่อนไข:\n" + "\n".join(error_details))
        comment_parts.append("💡 ข้อเสนอแนะเพื่อแก้ไข:\n" + "\n".join(fix_suggestions))
    
    # 2. ส่วนของ Missing Data (ลืมกรอกข้อมูล)
    if missing_status:
        # ถ้ามีส่วนแรกอยู่แล้ว ให้เว้นบรรทัด 1 ครั้งเพื่อความชัดเจน
        if status_parts:
            status_parts.append("\n" + "-"*30)
            comment_parts.append("\n" + "-"*30)
            
        status_parts.append("\n".join(missing_status))
        comment_parts.append("\n".join(missing_comment))

    # ถ้าไม่มีอะไรผิดเลย และไม่มีช่องว่าง
    if not status_parts:
        return "✅ ผ่านเงื่อนไขเบื้องต้น", "ข้อมูลถูกต้องและครบถ้วนตามเงื่อนไข"
    
    final_status = "\n".join(status_parts)
    final_comment = "\n".join(comment_parts)
    
    return final_status, final_comment

