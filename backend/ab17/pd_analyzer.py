import pandas as pd
import os
import re

class PDAnalyzer:
    def __init__(self, standards_dir):
        self.standards_dir = standards_dir
        self.standards_data = {}
        self.row_map = {}
        self._load_standards()

    def _load_standards(self):
        if not os.path.exists(self.standards_dir): return
        for filename in os.listdir(self.standards_dir):
            if filename.endswith(".xlsx") and len(filename) > 2 and filename[1] == ".":
                try:
                    full_path = os.path.join(self.standards_dir, filename)
                    # พิเศษสำหรับ e. (IEC60228_Table4)
                    if "IEC60228_Table4" in filename:
                        df = pd.read_excel(full_path, header=None)
                        std_map = {}
                        for _, row in df.iloc[7:].iterrows():
                            area_str = str(row[0]).replace(",", ".").replace(" ", "").strip()
                            area_match = re.search(r"(\d+\.?\d*)", area_str)
                            if area_match:
                                area_val = float(area_match.group(1))
                                min_wires = self.parse_value(row[3])
                                if min_wires > 0:
                                    std_map[area_val] = min_wires
                        self.standards_data['e_wires'] = std_map
                    
                    elif "Table C.2 IEC60228" in filename:
                        df = pd.read_excel(full_path, header=None)
                        f_map = []
                        for _, row in df.iloc[3:].iterrows():
                            area = self.parse_value(row[3])
                            mini = self.parse_value(row[4])
                            maxi = self.parse_value(row[5])
                            if area > 0:
                                f_map.append({"area": area, "min": mini, "max": maxi})
                        self.standards_data['f_diameters'] = f_map

                    elif "Table 4-4 IEC60288" in filename:
                        df = pd.read_excel(full_path, header=None)
                        h_map = []
                        last_v_range = ""
                        for _, row in df.iloc[4:].iterrows():
                            v_range = str(row[1]).strip()
                            if not v_range or v_range == 'nan':
                                v_range = last_v_range
                            else:
                                last_v_range = v_range
                                
                            a_range = str(row[2])
                            mini = self.parse_value(row[3])
                            maxi = self.parse_value(row[4])
                            if mini > 0:
                                h_map.append({"v_range": v_range, "a_range": a_range, "min": mini, "max": maxi})
                        self.standards_data['h_thickness'] = h_map

                    elif "Table 7-4" in filename:
                        df = pd.read_excel(full_path, header=None)
                        m_map = []
                        for _, row in df.iloc[4:].iterrows():
                            d_range = str(row[1])
                            mini_thick = self.parse_value(row[3])
                            m_map.append({"d_range": d_range, "min_thick": mini_thick})
                        self.standards_data['m_thickness'] = m_map

                    self.standards_data[filename[0].lower()] = pd.read_excel(full_path)
                except Exception as e:
                    print(f"Error loading {filename}: {e}")

    def parse_range(self, range_str, value):
        """เช็คว่า value อยู่ใน range_str หรือไม่ (เช่น '17.81 - 38.10', '0.700 or less', '25.01 and larger')"""
        if not range_str or range_str == 'nan': return False
        s = range_str.lower().replace(',', '').strip()
        
        # กรณี x - y
        match_between = re.search(r"(\d+\.?\d*)\s*-\s*(\d+\.?\d*)", s)
        if match_between:
            return float(match_between.group(1)) <= value <= float(match_between.group(2))
            
        # กรณี or less
        match_less = re.search(r"(\d+\.?\d*)\s*or\s*less", s)
        if match_less:
            return value <= float(match_less.group(1))
            
        # กรณี and larger
        match_larger = re.search(r"(\d+\.?\d*)\s*and\s*larger", s)
        if match_larger:
            return value >= float(match_larger.group(1))
            
        return False

    def find_voltage(self, df_pd, row_map):
        """หาค่าแรงดันไฟฟ้าจากข้อ d (Voltage class) โดยเฉพาะ"""
        if 'd' not in row_map: return 0.0
        
        start_idx = row_map['d']
        for offset in range(3): # ตรวจสอบแถว d และ 2 แถวถัดไป
            curr_idx = start_idx + offset
            if curr_idx >= len(df_pd): break
            row_str = " ".join(str(v) for v in df_pd.iloc[curr_idx].values if pd.notna(v)).lower()
            
            # ค้นหาตัวเลขที่อยู่ใกล้ kv
            match_kv = re.search(r"(\d+\.?\d*)\s*kv", row_str)
            if match_kv: return float(match_kv.group(1)) * 1000
            
            # fallback หา V
            match_v2 = re.search(r"(\d{2,})\s*v\b", row_str)
            if match_v2: return float(match_v2.group(1))
            
        return 0.0

    def parse_value(self, text):
        if pd.isna(text) or text == "": return -1.0
        if isinstance(text, (int, float)): return float(text)
        s = str(text).lower().replace(',', '.').replace(' ', '').strip()
        # หาตัวเลข (รองรับทศนิยม)
        match = re.search(r"(\d+\.\d+|\d+)", s)
        if match:
            val = float(match.group(1))
            # ข้ามเลขมาตรฐานหรือเลขปี
            if val in [60228, 60502, 2024, 2025, 2026, 2027]: return -1.0
            return val
        return -1.0

    def analyze_sections(self, df_pd):
        self.row_map = {}
        results = {}
        
        # 1. ค้นหาตำแหน่งหัวข้อ (Marker)
        for idx, row in df_pd.iterrows():
            row_str = " ".join(str(v) for v in row.values if pd.notna(v)).strip()
            match = re.search(r"^\s*([a-s])[\.\)]", row_str, re.I)
            if match: self.row_map[match.group(1).lower()] = idx

        # ฟังก์ชันดึงข้อมูลแบบกว้าง (กวาดทุกช่องในโซน)
        def get_zone_data(key, row_limit=2):
            key = key.lower()
            if key not in self.row_map: return {"val": 0.0, "text": "", "all_cells": [], "all_nums": []}
            
            start_idx = self.row_map[key]
            all_cells = []
            found_nums = []
            
            for offset in range(row_limit):
                curr_idx = start_idx + offset
                if curr_idx >= len(df_pd): break
                row_data = df_pd.iloc[curr_idx]
                
                for i, v in enumerate(row_data.values):
                    if pd.isna(v): continue
                    s_val = str(v).strip()
                    if not s_val: continue
                    
                    # ข้าม Item No. (เช่น e.)
                    if i == 0 and offset == 0: continue
                    
                    all_cells.append(s_val.lower())
                    num = self.parse_value(v)
                    if num >= 0: found_nums.append(num)
            
            return {
                "val": found_nums[0] if found_nums else 0.0,
                "text": " ".join(all_cells),
                "all_cells": all_cells,
                "all_nums": found_nums
            }

        # --- วิเคราะห์รายข้อ (v3.0 - Flexible Keyword Matching) ---
        
        def check_keywords(text, keywords, logic="or"):
            text = text.lower()
            if logic == "or":
                return any(kw.lower() in text for kw in keywords)
            else: # and logic
                return all(kw.lower() in text for kw in keywords)

        def has_copper_keyword(text):
            lowered = str(text or "").lower()
            if "copper" in lowered or "ทองแดง" in lowered:
                return True
            # Match chemical symbol as a standalone token, not inside words like
            # "calculated" which previously caused false positives.
            return re.search(r"(?<![a-z0-9])cu(?![a-z0-9])", lowered) is not None

        # 0. ดึงข้อมูลพื้นที่และแรงดันก่อน (Common Data)
        e_data = get_zone_data('e', row_limit=4)
        area_val = e_data['val']
        voltage_val = self.find_voltage(df_pd, self.row_map)

        # f (Diameter)
        f_data = get_zone_data('f')
        f_val = f_data['val']
        f_err = []
        if area_val > 0 and 'f_diameters' in self.standards_data:
            # 1. หาแถวในตาราง f ที่มี area ใกล้เคียง area_val (error 10%)
            matched_f = None
            for row_f in self.standards_data['f_diameters']:
                if abs(row_f['area'] - area_val) <= (row_f['area'] * 0.10):
                    matched_f = row_f
                    break
            
            if matched_f:
                # 2. เช็ค f_val ว่าอยู่ในช่วง [min, maxi] (ยอมรับ error 5%)
                mini, maxi = matched_f['min'], matched_f['max']
                if f_val > 0:
                    if f_val < (mini * 0.95) or f_val > (maxi * 1.05):
                        f_err.append(f"Diameter {f_val} ไม่อยู่ในช่วงมาตรฐาน {mini}-{maxi} mm")
            else:
                f_err.append(f"ไม่พบพื้นที่หน้าตัด {area_val} ในตารางมาตรฐาน f")
        
        results['f'] = {"status": "ผิด" if f_err else "ผ่าน", "comment": "; ".join(f_err) or "ผ่าน"}

        # e. Calculated Area & Material & Number of Wires
        e_err = []
        
        # 1. เช็คพื้นที่หน้าตัด
        if 0 < area_val < 1.49:
            e_err.append(f"พื้นที่ {area_val} น้อยกว่า 1.5")
        
        # 2. เช็คจำนวน Wires (IEC 60228 Table 4)
        if area_val > 0 and 'e_wires' in self.standards_data:
            # หาค่าที่ใกล้เคียงที่สุดในตารางมาตรฐาน (v3.2: ใช้ค่ามาตรฐานที่น้อยกว่าหรือเท่ากับพื้นที่ที่คำนวณได้)
            std_wires = -1
            sorted_std_areas = sorted(self.standards_data['e_wires'].keys(), reverse=True)
            for std_area in sorted_std_areas:
                if area_val >= std_area: # เช่น 35.7 >= 35.0
                    std_wires = self.standards_data['e_wires'][std_area]
                    break
            
            if std_wires > 0:
                bidder_wires = -1
                start_idx_e = self.row_map.get('e')
                
                # พยายามหาจำนวน wires โดยดูจากคำอธิบายข้างๆ (v3.1)
                for offset in range(4):
                    if start_idx_e is None: break
                    curr_idx = start_idx_e + offset
                    if curr_idx >= len(df_pd): break
                    row_data = df_pd.iloc[curr_idx]
                    row_text = " ".join(str(v).lower() for v in row_data.values if pd.notna(v))
                    
                    if check_keywords(row_text, ["no. of wire", "number of wire", "no of wire"]):
                        # หาตัวเลขในแถวนี้ที่ทำหน้าที่เป็นจำนวนสาย
                        for v in row_data.values:
                            num = self.parse_value(v)
                            if num > 1 and num != area_val:
                                bidder_wires = num
                                break
                    if bidder_wires > 0: break

                if bidder_wires > 0:
                    if bidder_wires < std_wires:
                        e_err.append(f"จำนวนสาย {int(bidder_wires)} น้อยกว่ามาตรฐาน (ขั้นต่ำ {int(std_wires)})")
                elif area_val >= 1.5:
                    e_err.append(f"ไม่พบข้อมูลจำนวนสาย (ต้องไม่น้อยกว่า {int(std_wires)})")
        
        # 3. ตรวจวัสดุ (Flexible)
        if not has_copper_keyword(e_data['text']):
            e_err.append("ไม่พบวัสดุทองแดง (Copper)")
            
        results['e'] = {"status": "ผิด" if e_err else "ผ่าน", "comment": "; ".join(e_err) or "ผ่าน"}

        # g (Conductor Screen) - Flexible
        g_data = get_zone_data('g')
        g_text = g_data['text']
        g_err = []
        
        # คีย์เวิร์ดสำคัญสำหรับ g
        semi_cond = check_keywords(g_text, ["semi", "conduct"])
        thermo = check_keywords(g_text, ["thermoset", "xlpe"]) # เพิ่ม XLPE เป็นทางเลือกของ thermosetting
        tape = check_keywords(g_text, ["tape"])

        if "1." in g_text and "2." in g_text:
            # กรณีระบุ 2 อย่าง
            if not (semi_cond and (tape or thermo)): # ยืดหยุ่น: ขอให้มี semi และ (tape หรือ thermo)
                g_err.append("ต้องมี Semi-conducting tape หรือวัสดุที่กำหนด")
        else:
            # กรณีระบุอย่างเดียว
            if not (semi_cond):
                g_err.append("ต้องมีวัสดุ Semi-conducting")
            if not (thermo or tape):
                g_err.append("ต้องเป็นวัสดุ Thermosetting หรือ Tape")
        
        results['g'] = {"status": "ผิด" if g_err else "ผ่าน", "comment": "; ".join(g_err) or "ผ่าน"}
        results['g_thick'] = g_data['val']

        # i (Insulation Screen) - Flexible
        i_data = get_zone_data('i')
        # ยืดหยุ่นมากขึ้นสำหรับ i: ขอแค่มีคำที่บ่งบอกว่าเป็นกึ่งตัวนำ หรือ XLPE ที่เป็น semi
        is_semi = check_keywords(i_data['text'], ["semi", "conduct"])
        is_material = check_keywords(i_data['text'], ["tape", "เทป", "xlpe", "thermoset"])
        
        if not (is_semi or is_material):
            results['i'] = {"status": "ผิด", "comment": "วัสดุต้องเป็น semi-conductive, XLPE หรือ tape"}
        else: results['i'] = {"status": "ผ่าน", "comment": "ผ่าน"}
        results['i_thick'] = i_data['val']

        # h (Insulation) - Flexible
        h_data = get_zone_data('h')
        h_val = h_data['val']
        h_err = []
        is_xlpe = check_keywords(h_data['text'], ["xlpe"])
        if not is_xlpe: h_err.append("ต้องเป็น XLPE")
        
        print(f"\n[DEBUG h] Voltage found: {voltage_val}")
        print(f"[DEBUG h] h_val provided: {h_val}")
        
        # เช็คความหนาตามตาราง h (v4.2: แจ้งเตือนหากไม่พบแรงดัน)
        if 'h_thickness' in self.standards_data:
            if voltage_val > 0:
                matched_h = None
                for row_h in self.standards_data['h_thickness']:
                    if self.parse_range(row_h['v_range'], voltage_val):
                        matched_h = row_h
                        print(f"[DEBUG h] Matched row: {row_h}")
                        break
                
                if matched_h:
                    mini, maxi = matched_h['min'], matched_h['max']
                    if h_val > 0:
                        if h_val > maxi:
                            h_err.append(f"ความหนา {h_val} เกินมาตรฐานสูงสุด {maxi} mm")
                        if h_val < mini:
                            h_err.append(f"ความหนา {h_val} น้อยกว่ามาตรฐานขั้นต่ำ {mini} mm")
                else:
                    h_err.append(f"ไม่พบช่วงแรงดัน {voltage_val} V ในตารางมาตรฐาน h")
            else:
                h_err.append("⚠️ ไม่พบข้อมูลแรงดันไฟฟ้า (kV) ในไฟล์ จึงไม่สามารถตรวจความหนาได้")
        else:
            print("[DEBUG h] 'h_thickness' table NOT found in standards_data!")

        results['h'] = {"status": "ผิด" if h_err else "ผ่าน", "comment": "; ".join(h_err) or "ผ่าน"}
        results['h_thick'] = h_val

        # k (Metal Screen) - Flexible
        k = get_zone_data('k')
        k_err = []
        if 0 < k['val'] < 0.1: k_err.append(f"ความหนา {k['val']} mm น้อยกว่า 0.1 mm")
        if not has_copper_keyword(k['text']): 
            k_err.append("ต้องเป็นวัสดุ Copper")
        results['k'] = {"status": "ผิด" if k_err else "ผ่าน", "comment": "; ".join(k_err) or "ผ่าน"}
        results['k_thick'] = k['val']

        # l (Cushion) - Flexible
        l = get_zone_data('l')
        l_err = []
        if 0 < l['val'] < 0.029:
            l_err.append(f"ความหนา {l['val']} mm น้อยกว่า 0.03 mm")
        
        # เช็คคีย์เวิร์ด Polyester และ Non-woven (ไม่สนลำดับหรือขีด)
        has_polyester = check_keywords(l['text'], ["polyester"])
        has_nonwoven = check_keywords(l['text'], ["non", "woven"]) # เช็คว่ามีทั้ง non และ woven
        
        if not (has_polyester and has_nonwoven):
            l_err.append("วัสดุต้องเป็น Polyester non-woven tape")
            
        results['l'] = {"status": "ผิด" if l_err else "ผ่าน", "comment": "; ".join(l_err) or "ผ่าน"}
        results['l_thick'] = l['val']

        # m (Oversheath) - Flexible
        m_data = get_zone_data('m')
        m_val = m_data['val']
        m_err = []
        is_pvc = check_keywords(m_data['text'], ["pvc"])
        if not is_pvc: m_err.append("ต้องเป็น PVC")
        
        # เช็คความหนาตามตาราง m
        d_under_jacket = f_val + (2 * (results['g_thick'] + results['h_thick'] + results['i_thick'] + results['k_thick'] + results['l_thick']))
        if d_under_jacket > 0 and 'm_thickness' in self.standards_data:
            matched_m = None
            for row_m in self.standards_data['m_thickness']:
                if self.parse_range(row_m['d_range'], d_under_jacket):
                    matched_m = row_m
                    break
            
            if matched_m:
                min_m = matched_m['min_thick']
                if m_val > 0 and m_val < min_m:
                    m_err.append(f"ความหนา {m_val} น้อยกว่ามาตรฐานขั้นต่ำ {min_m} mm")

        results['m'] = {"status": "ผิด" if m_err else "ผ่าน", "comment": "; ".join(m_err) or "ผ่าน"}
        results['m_thick'] = m_val

        # --- Formulas (j & n) ---
        g_t, h_t, i_t = results['g_thick'], results['h_thick'], results['i_thick']
        k_t, l_t, m_t = results['k_thick'], results['l_thick'], results['m_thick']

        j = get_zone_data('j')
        if f_val > 0 and j['val'] > 0:
            exp_j = f_val + (2 * (g_t + h_t + i_t))
            # ปรับ Tolerance เป็น 10% ตามคำขอของผู้ใช้
            if abs(j['val'] - exp_j) > (exp_j * 0.10): results['j'] = {"status": "ผิด", "comment": f"ค่า j ไม่ตรงสูตร (ควรประมาณ {exp_j:.2f})"}
            else: results['j'] = {"status": "ผ่าน", "comment": "ผ่าน"}
        else: results['j'] = {"status": "ผ่าน", "comment": "ผ่าน"}

        n = get_zone_data('n')
        if f_val > 0 and n['val'] > 0:
            exp_n = f_val + (2 * (g_t + h_t + i_t + k_t + l_t + m_t))
            # ปรับ Tolerance เป็น 10% ตามคำขอของผู้ใช้
            if abs(n['val'] - exp_n) > (exp_n * 0.10): results['n'] = {"status": "ผิด", "comment": f"ค่า n ไม่ตรงสูตร (ควรประมาณ {exp_n:.2f})"}
            else: results['n'] = {"status": "ผ่าน", "comment": "ผ่าน"}
        else: results['n'] = {"status": "ผ่าน", "comment": "ผ่าน"}

        for key in "abcdopqrs":
            if key not in results: results[key] = {"status": "ผ่าน", "comment": "ผ่าน"}
            
        return results
