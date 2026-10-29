#!/usr/bin/env python3
import os
import re
import sys
import json
import datetime
from pathlib import Path
import openpyxl

TEXT_FIXES = [
    (r"Dipheyl", "Diphenyl"),
    (r"\bFluorescamin\b", "Fluorescamine"),
    (r"PRimeScript RT reagent kKit", "PrimeScript RT reagent Kit"),
    (r"\blgG\b", "IgG"),
    (r"CD806", "CD86"),
    (r"cellROX", "CellROX"),
    (r"Mitotracker", "MitoTracker"),
    (r"뒷편", "뒤편"),
    (r"공구 서랍(\d)", r"공구 서랍 \1"),
    (r"Weight Dish", "Weighing Dish"),
    (r"밧드", "바트"),
    (r"칼블럭", "칼브럭"),
    (r"에포치", "에포크"),
    (r"폐쇄기 \(paper shredder\)", "파쇄기 (Paper Shredder)"),
    (r"폐쇄기", "파쇄기"),
    (r"pro analysi\b", "pro analysis"),
    (r"Alexa fluor", "Alexa Fluor"),
    (r"donkey anti-rabbit", "Donkey anti-Rabbit"),
    (r"\bclostridium histolyticum", "Clostridium histolyticum"),
    (r"Cellbanker1", "Cellbanker 1"),
    (r"methylether", "methyl ether"),
    (r"CD 11b", "CD11b"),
    (r"Molecularporous", "Molecular/Porous"),
    (r"ERGONINIC", "ERGONOMIC"),
    (r"\bAPOCH\b", "EPOCH"),
    (r"tektronics", "Tektronix"),
    (r"\bphcbi\b", "PHCbi"),
    (r"전동 파이펫", "전동 피펫"),
    (r"simulation mod\b", "simulation module"),
    (r"Stabilization Solutionwi\b", "Stabilization Solution"),
    (r"일렉트로스피닝 장치 \(Spinning Device\)", "전기방사 장치 (Electrospinning Device)"),
    (r"^neurolux$", "Neurolux"),
    (r"neurolux inc", "Neurolux Inc."),
    (r"^generator \(", "Generator ("),
    (r"(?<=\d) uL\b", " µL"),
]

MANUFACTURER_FIXES = {
    "Guangdong Victoy Biotech": "Guangdong Victory Biotech",
    "GoodFellow": "Goodfellow",
    "ANASPEC": "AnaSpec",
    "Welgene": "WELGENE",
    "dojindo": "DojinDo",
    "abcam": "Abcam",
    "阿拉丁(Aladdin)": "Aladdin",
}


def apply_text_fixes(text):
    for pattern, replacement in TEXT_FIXES:
        text = re.sub(pattern, replacement, text)
    return text


def normalize_capacity_text(text):
    """Unify mu characters and spacing/casing of plain '<number><unit>' capacities."""
    text = text.replace("μ", "µ")
    text = re.sub(r"(?<=\d) ?ug\b", " µg", text)
    match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*(mg|kg|g|ml|l|ul|µl)", text, re.IGNORECASE)
    if not match:
        return text
    unit = {"ml": "mL", "l": "L", "ul": "µL", "µl": "µL"}.get(match.group(2).lower(), match.group(2).lower())
    return f"{match.group(1)} {unit}"


def get_bilingual_name(name_val):
    if name_val is None:
        return ""
    name = str(name_val).strip()
    name_map = {
        "Minisaw": "미니톱 (Minisaw)",
        "저울": "저울 (Scale)",
        "미세저울": "미세저울 (Microscale)",
        "가열교반기": "가열교반기 (Hot Plate Stirrer)",
        "pH/이온 측정기/ 비색계": "pH/이온 측정기/비색계 (pH/Ion Meter & Colorimeter)",
        "접촉각": "접촉각 측정기 (Contact Angle Meter)",
        "Spin Coater": "스핀 코터 (Spin Coater)",
        "UTM": "만능 재료 시험기 (UTM)",
        "incubator": "인큐베이터 (Incubator)",
        "형광 단백질 광원": "형광 단백질 광원 (Fluorescence Light Source)",
        "Mimics S/W": "미믹스 소프트웨어 (Mimics S/W)",
        "광조사기": "광조사기 (Light Curing Unit)",
        "멸균기": "고압멸균기 (Autoclave)",
        "Vacuum oven": "진공 오븐 (Vacuum Oven)",
        "Epoch": "에포크 마이크로플레이트 리더 (Epoch Microplate Reader)",
        "초음파 세척기": "초음파 세척기 (Ultrasonic Cleaner)",
        "Thermo CO2 incubator": "써모 CO2 인큐베이터 (Thermo CO2 Incubator)",
        "질소탱크": "질소탱크 (Nitrogen Tank)",
        "Real-Time PCR": "실시간 유전자 증폭기 (Real-Time PCR)",
        "Centrifuge": "원심분리기 (Centrifuge)",
        "thermal cycler": "열 순환기 (Thermal Cycler)",
        "phcbi CO2 incubator": "phcbi CO2 인큐베이터 (phcbi CO2 Incubator)",
        "uniz": "유니즈 3D 프린터 (UNIZ 3D Printer)",
        "질소 경화기": "질소 경화기 (Nitrogen Curing Chamber)",
        "Spining": "일렉트로스피닝 장치 (Spinning Device)",
        "회전증발 농축기": "회전증발 농축기 (Rotary Evaporator)",
        "ultrasonic homogenizer probes": "초음파 균질기 프로브 (Ultrasonic Homogenizer Probes)",
        "UV 경화기 2개": "UV 경화기 - 2개 (UV Curing Units - 2pcs)",
        "큐레이": "큐레이 구강카메라 (Q-ray Oral Camera)",
        "electrospinning": "전기방사 장치 (Electrospinning)",
        "구강 스캐너": "구강 스캐너 (Intraoral Scanner)",
        "FT-IR": "적외선 분광 광도계 (FT-IR)",
        "연마기": "연마기 (Polishing Machine)",
        "자동 제습 보관함": "자동 제습 보관함 (Auto-Desiccating Cabinet)",
        "진공 데시게이터": "진공 데시게이터 (Vacuum Desiccator)",
        "EVOS M5000 이미징시스템": "EVOS M5000 이미징 시스템 (EVOS M5000 Imaging System)",
        "볼텍스 믹서": "볼텍스 믹서 (Vortex Mixer)",
        "냉장고 (세균, 공용)": "공용 세균 냉장고 (Shared Refrigerator)",
        "사진촬영": "암실 사진 촬영기 (Photo Documentation System)",
        "전동 파이펫, 파이펫": "전동 파이펫 및 피펫 세트 (Pipette Aid & Pipette Set)",
        "현미경": "광학 현미경 (Microscope)",
        "교정용 Vac Plus": "교정용 백플러스 (Calibration Vac Plus)",
        "PC 본체": "데스크탑 PC 본체 (PC Desktop Tower)",
        "폐쇄기": "폐쇄기 (paper shredder)",
        "프린터": "사무용 프린터 (Office Printer)",
        "노트북": "노트북 컴퓨터 (Laptop Computer)",
        "마이크로칩": "마이크로플루이딕 마이크로칩 (Microfluidic Microchip)",
        "펌프 장비": "진공 펌프 (Vacuum Pump)",
        "UV 멸균": "UV 멸균기 (UV Sterilizer)"
    }
    return name_map.get(name, name)

def sync():
    root = Path(__file__).resolve().parents[1]
    html_path = root / "index.html"
    manager_path = root / "담당자.xlsx"
    equipment_files = sorted(
        [f for f in root.glob("*장비목록*.xlsx") if f.is_file() and not f.name.startswith("~$")],
        key=lambda x: x.name
    )
    
    if not equipment_files:
        print("Error: No equipment list workbook (*장비목록*.xlsx) found.")
        return False
    excel_path = equipment_files[-1]

    # The equipment list has no manager column, so managers come from 담당자.xlsx (keyed by No.).
    managers = {}
    if manager_path.exists():
        try:
            wb_mgr = openpyxl.load_workbook(manager_path, read_only=True, data_only=True)
            mgr_sheet = next((n for n in wb_mgr.sheetnames if n.strip() == "장비 목록 (전체)"), None)
            if mgr_sheet is not None:
                for row in wb_mgr[mgr_sheet].iter_rows(min_row=4, values_only=True):
                    if len(row) < 7 or row[1] is None:
                        continue
                    try:
                        mgr_no = int(row[1])
                    except (ValueError, TypeError):
                        continue
                    mgr_name = str(row[5]).strip() if row[5] is not None else ""
                    managers[mgr_no] = ("" if mgr_name == "미지정" else mgr_name,
                                        str(row[6]).strip() if row[6] is not None else "")
            wb_mgr.close()
        except Exception as e:
            print(f"Warning: Could not read managers from {manager_path.name}: {e}")
    else:
        print(f"Warning: {manager_path.name} not found; managers will be empty.")
        
    print(f"Reading general equipment from {excel_path.name}...")
    try:
        wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        inventory_data = []
        
        for r_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if len(row) < 3:
                continue
                
            no_val = row[0]
            name_val = row[1]
            loc_val = row[2]
            mgr_val, remarks_val = managers.get(no_val if isinstance(no_val, int) else None, ("", ""))
            
            if no_val is None:
                continue
                
            try:
                no = int(no_val)
            except (ValueError, TypeError):
                continue
                
            if not name_val:
                continue
                
            loc = str(loc_val).strip() if loc_val is not None else ""
            if loc == "미지정":
                loc = ""
                
            mgr = str(mgr_val).strip() if mgr_val is not None else ""
            if mgr == "미지정":
                mgr = ""
                
            remarks = str(remarks_val).strip() if remarks_val is not None else ""
            
            inventory_item = {
                "no": no,
                "name": get_bilingual_name(str(name_val).replace("pH_이온", "pH/이온")),
                "loc": loc,
                "mgr": mgr,
                "remarks": remarks
            }
            if no == 43 or "폐쇄기" in str(name_val):
                inventory_item["loc"] = "대신 building"
                inventory_item["icon"] = "rat"
            inventory_data.append(inventory_item)
            
        wb.close()
        print(f"Successfully parsed {len(inventory_data)} general equipment entries.")
        
    except Exception as e:
        print(f"An error occurred while reading {excel_path.name}: {e}")
        return False

    # Now read pipette and PC from archive
    pipette_data = []
    pc_data = []
    
    archive_dir = root / "archive"
    archive_files = []
    if archive_dir.exists():
        archive_files = sorted(
            [f for f in archive_dir.glob("*_장비 보유 현황.xlsx") if f.is_file()],
            key=lambda x: x.name
        )
        
    if not archive_files:
        print("Warning: No matching *_장비 보유 현황.xlsx found in archive/ folder. Skipping Pipette and PC sync.")
    else:
        latest_archive = archive_files[-1]
        print(f"Reading Pipette & PC from latest archive: {latest_archive.name}...")
        try:
            wb_arc = openpyxl.load_workbook(latest_archive, read_only=True, data_only=True)
            
            # Parse Pipette
            if "Pipette" in wb_arc.sheetnames:
                ws_pip = wb_arc["Pipette"]
                # Headers are on row 2, data starts on row 3
                for row in ws_pip.iter_rows(min_row=3, values_only=True):
                    if not row or len(row) < 5:
                        continue
                    no_val = row[0]
                    name_val = row[1]
                    spec_val = row[2]
                    loc_val = row[3]
                    brand_val = row[4]
                    calib_val = row[5] if len(row) > 5 else None
                    broken_val = row[6] if len(row) > 6 else None
                    remarks_val = row[7] if len(row) > 7 else None
                    
                    if no_val is None:
                        continue
                    try:
                        no = int(no_val)
                    except (ValueError, TypeError):
                        continue
                    if not name_val:
                        continue
                        
                    loc = str(loc_val).strip() if loc_val is not None else ""
                    if loc == "미지정":
                        loc = ""
                    brand = str(brand_val).strip() if brand_val is not None else ""
                    calib = str(calib_val).strip() if calib_val is not None else ""
                    broken = str(broken_val).strip() if broken_val is not None else ""
                    remarks = str(remarks_val).strip() if remarks_val is not None else ""
                    
                    pipette_data.append({
                        "no": no,
                        "name": str(name_val).strip(),
                        "spec": spec_val or "",
                        "loc": loc,
                        "brand": brand,
                        "calib": calib,
                        "broken": broken,
                        "remarks": remarks
                    })
                print(f"Successfully parsed {len(pipette_data)} pipette entries.")
            else:
                print("Warning: 'Pipette' sheet not found in the archive Excel.")
                
            # Parse PC
            if "PC" in wb_arc.sheetnames:
                ws_pc = wb_arc["PC"]
                # Headers are on row 2, data starts on row 3
                for row in ws_pc.iter_rows(min_row=3, values_only=True):
                    if not row or len(row) < 5:
                        continue
                    no_val = row[0]
                    kind_val = row[1]
                    label_val = row[2]
                    loc_val = row[3]
                    user_val = row[4]
                    brand_val = row[5]
                    photo_val = row[6] if len(row) > 6 else None
                    sticker_val = row[7] if len(row) > 7 else None
                    
                    if no_val is None:
                        continue
                    try:
                        no = int(no_val)
                    except (ValueError, TypeError):
                        continue
                    if not kind_val:
                        continue
                        
                    loc = str(loc_val).strip() if loc_val is not None else ""
                    if loc == "미지정":
                        loc = ""
                    kind = str(kind_val).strip()
                    label = str(label_val).strip() if label_val is not None else ""
                    user = str(user_val).strip() if user_val is not None else ""
                    brand = str(brand_val).strip() if brand_val is not None else ""
                    photo = str(photo_val).strip() if photo_val is not None else ""
                    sticker = str(sticker_val).strip() if sticker_val is not None else ""
                    
                    pc_data.append({
                        "no": no,
                        "kind": kind,
                        "label": label,
                        "loc": loc,
                        "user": user,
                        "brand": brand,
                        "photo": photo,
                        "sticker": sticker
                    })
                print(f"Successfully parsed {len(pc_data)} PC entries.")
            else:
                print("Warning: 'PC' sheet not found in the archive Excel.")
                
            wb_arc.close()
        except Exception as e:
            print(f"An error occurred while reading the archive Excel: {e}")
            # Keep empty lists to avoid breaking the script
            
    # Read the dedicated animal-lab inventory from its source workbook.
    animal_lab_data = []
    animal_lab_path = root / "archive" / "동물실험실 관련 물품 리스트.xlsx"
    if animal_lab_path.exists():
        try:
            wb_animal = openpyxl.load_workbook(animal_lab_path, read_only=True, data_only=True)
            ws_animal = wb_animal[wb_animal.sheetnames[0]]
            for row in ws_animal.iter_rows(min_row=3, values_only=True):
                code_val = row[2] if len(row) > 2 else None
                name_val = row[3] if len(row) > 3 else None
                if not code_val or not name_val:
                    continue
                no = len(animal_lab_data) + 1
                animal_lab_data.append({
                    "no": no,
                    "code": str(code_val).strip(),
                    "name": str(name_val).strip(),
                    "photo": f"animal_lab_{no}",
                    "loc": "ABMRC animal-research-lab B3/B4",
                    "mgr": "김재형, 정수미"
                })
            wb_animal.close()
            print(f"Successfully parsed {len(animal_lab_data)} animal-lab entries.")
        except Exception as e:
            print(f"Warning: Could not read animal-lab workbook: {e}")
    else:
        print(f"Warning: Animal-lab workbook not found at {animal_lab_path}.")

    # Read reagent list workbook
    reagent_data = []
    reagent_files = sorted(
        [f for f in root.glob("*시약_물품_리스트*.xlsx") if not f.name.startswith("~$") and f.is_file()],
        key=lambda x: x.name
    )
    if not reagent_files:
        reagent_files = sorted(
            [f for f in root.glob("시약_리스트_*.xlsx") if not f.name.startswith("~$") and f.is_file()],
            key=lambda x: x.name
        )
    if not reagent_files:
        # Fallback to archive folder if exists
        if archive_dir.exists():
            reagent_files = sorted(
                [f for f in archive_dir.glob("시약_리스트_*.xlsx") if not f.name.startswith("~$") and f.is_file()],
                key=lambda x: x.name
            )

    if reagent_files:
        latest_reagent = reagent_files[-1]
        print(f"Reading Reagent inventory from {latest_reagent.name}...")
        try:
            wb_reagent = openpyxl.load_workbook(latest_reagent, read_only=True, data_only=True)
            def normalize(header):
                return re.sub(r"\s+", "", str(header).strip()).lower()

            # Sheets with these headers are read; the 시약/재료/Buffers sheets are reagents
            # and the 장비 sheet (장비명 column) holds equipment. Supply sheets are ignored.
            def sheet_rows():
                for ws_reagent in wb_reagent.worksheets:
                    header_row = next(ws_reagent.iter_rows(min_row=1, max_row=1, values_only=True), ())
                    header_map = {normalize(h): i for i, h in enumerate(header_row) if h is not None}
                    name_header = next((h for h in ("시약명", "장비명") if normalize(h) in header_map), None)
                    required = {"보관 위치", "구역", "세부", "제조사", "Cat.No (제품번호)", "CAS No.", "용량", "수량", "비고"}
                    if name_header is None or any(normalize(h) not in header_map for h in required):
                        print(f"Skipping sheet '{ws_reagent.title}' (not a reagent/equipment sheet).")
                        continue
                    for row in ws_reagent.iter_rows(min_row=2, values_only=True):
                        if row and any(c is not None for c in row):
                            yield ws_reagent.title, header_map, name_header, row

            type_labels = {"시약": "시약", "재료": "재료", "buffers": "Buffers"}
            for sheet_title, header_map, name_header, row in sheet_rows():
                def cell_value(row, *header_names):
                    for header_name in header_names:
                        index = header_map.get(normalize(header_name))
                        if index is not None and index < len(row):
                            return row[index]
                    return None

                category_val = "장비" if name_header == "장비명" else cell_value(row, "구분")
                loc_val = cell_value(row, "보관 위치")
                zone_val = cell_value(row, "구역")
                folder_val = cell_value(row, "세부")
                name_val = cell_value(row, name_header)
                mfr_val = cell_value(row, "제조사")
                cat_val = cell_value(row, "Cat.No (제품번호)")
                cas_val = cell_value(row, "CAS No.")
                capacity_val = cell_value(row, "용량")
                qty_val = cell_value(row, "수량")
                remarks_val = cell_value(row, "비고")

                if str(category_val).strip() == "장비":
                    equipment_name = str(name_val).strip() if name_val is not None else ""
                    existing = next((item for item in inventory_data if item["name"] == equipment_name), None)
                    equipment_location = str(loc_val).strip() if loc_val is not None else ""
                    equipment_zone = str(zone_val).strip() if zone_val is not None else ""
                    equipment_mfr = str(mfr_val).strip() if mfr_val is not None else ""
                    equipment_qty = str(qty_val).strip() if qty_val is not None else ""
                    equipment_note = str(remarks_val).strip() if remarks_val is not None else ""
                    equipment_remarks = [
                        value for value in (
                            f"보관 구역: {equipment_zone}" if equipment_zone not in ("", "-") else "",
                            f"제조사: {equipment_mfr}" if equipment_mfr else "",
                            f"수량: {equipment_qty}" if equipment_qty else "",
                            equipment_note
                        ) if value
                    ]
                    if existing is not None:
                        if not existing["remarks"]:
                            existing["remarks"] = "; ".join(equipment_remarks)
                        continue
                    inventory_data.append({
                        "no": len(inventory_data) + 1,
                        "name": equipment_name,
                        "loc": equipment_location,
                        "mgr": "",
                        "remarks": "; ".join(equipment_remarks)
                    })
                    continue

                name = str(name_val).strip() if name_val is not None else ""
                if not name:
                    continue

                loc = str(loc_val).strip() if loc_val is not None else ""
                zone = str(zone_val).strip() if zone_val is not None else ""
                folder = str(folder_val).strip() if folder_val is not None else ""
                mfr = str(mfr_val).strip() if mfr_val is not None else ""
                cat = str(cat_val).strip() if cat_val is not None else ""
                cas = str(cas_val).strip() if cas_val is not None else ""
                capacity = str(capacity_val).strip() if capacity_val is not None else ""
                qty = str(qty_val).strip() if qty_val is not None else ""
                if qty.endswith(".0"):
                    qty = qty[:-2]
                remarks = str(remarks_val).strip() if remarks_val is not None else ""

                reagent_data.append({
                    "no": len(reagent_data) + 1,
                    "type": type_labels.get(sheet_title.strip().lower(), sheet_title.strip()),
                    "loc": loc,
                    "zone": zone,
                    "folder": folder,
                    "name": name,
                    "mfr": mfr,
                    "cat": cat,
                    "cas": cas,
                    "capacity": capacity,
                    "qty": qty,
                    "remarks": remarks
                })
            # Supply sheets ("<호실>-물품 및 소모품" and "<호실> Mold List") are listed as
            # stock items alongside the reagents so they appear in the same directory.
            def clean(value):
                return str(value).strip() if value is not None else ""

            for ws_supply in wb_reagent.worksheets:
                room_match = re.match(r"(\d+호)", ws_supply.title.strip())
                if not room_match:
                    continue
                room = room_match.group(1)
                is_mold = "mold" in ws_supply.title.lower()
                if not is_mold and "물품" not in ws_supply.title:
                    continue
                for row in ws_supply.iter_rows(min_row=3 if is_mold else 2, values_only=True):
                    if not row or all(c is None for c in row):
                        continue
                    if is_mold:
                        spec, _image, count = (list(row) + [None] * 3)[:3]
                        if not clean(spec):
                            continue
                        entry = {"zone": "Mold", "folder": "-", "name": f"Mold {clean(spec)}",
                                 "mfr": "", "capacity": clean(spec), "qty": clean(count), "remarks": ""}
                    else:
                        category, detail, item, brand, spec, qty, note = (list(row) + [None] * 7)[:7]
                        if not clean(item):
                            continue
                        entry = {"zone": clean(category) or "-", "folder": clean(detail) or "-",
                                 "name": clean(item), "mfr": clean(brand), "capacity": clean(spec),
                                 "qty": clean(qty), "remarks": clean(note)}
                    if entry["qty"].endswith(".0"):
                        entry["qty"] = entry["qty"][:-2]
                    reagent_data.append({
                        "no": len(reagent_data) + 1, "type": "소모품", "loc": room, "zone": entry["zone"],
                        "folder": entry["folder"], "name": entry["name"], "mfr": entry["mfr"],
                        "cat": "", "cas": "", "capacity": entry["capacity"], "qty": entry["qty"],
                        "remarks": entry["remarks"]
                    })
            wb_reagent.close()
            print(f"Successfully parsed {len(reagent_data)} reagent/supply entries.")
        except Exception as e:
            print(f"Warning: Could not read reagent workbook: {e}")
    else:
        print("Warning: No reagent workbook (시약_리스트_*.xlsx) found.")

    # Write to index.html
    for item in reagent_data:
        for field in ("name", "zone", "folder", "remarks"):
            item[field] = apply_text_fixes(item[field])
        item["mfr"] = MANUFACTURER_FIXES.get(item["mfr"], item["mfr"])
        item["capacity"] = normalize_capacity_text(item["capacity"])
    for dataset in (inventory_data, pipette_data, pc_data, animal_lab_data):
        for item in dataset:
            for field, value in item.items():
                if isinstance(value, str) and field not in ("image", "photo", "icon"):
                    item[field] = apply_text_fixes(value)
    for item in inventory_data:
        if "파쇄기" in item["name"]:
            item["icon"] = "rat"

    # Write to index.html
    if not html_path.exists():
        print(f"Error: index.html not found at {html_path}")
        return False
        
    try:
        html_content = html_path.read_text(encoding="utf-8")
        
        # Helper to format data
        def format_json(data_list):
            formatted_json = "[\n"
            for i, item in enumerate(data_list):
                comma = "," if i < len(data_list) - 1 else ""
                item_json = json.dumps(item, ensure_ascii=False)
                formatted_json += f"      {item_json}{comma}\n"
            formatted_json += "    ]"
            return formatted_json
            
        # Replace inventoryData
        formatted_inventory = format_json(inventory_data)
        pattern_inv = r"(const inventoryData = )\[[\s\S]*?\];"
        replacement_inv = f"\\1{formatted_inventory};"
        html_content, count_inv = re.subn(pattern_inv, replacement_inv, html_content)
        if count_inv == 0:
            print("Error: Could not locate 'const inventoryData = [...];' in index.html.")
            return False
            
        # Replace pipetteData
        formatted_pipette = format_json(pipette_data)
        pattern_pip = r"(const pipetteData = )\[[\s\S]*?\];"
        replacement_pip = f"\\1{formatted_pipette};"
        html_content, count_pip = re.subn(pattern_pip, replacement_pip, html_content)
        if count_pip == 0:
            print("Warning: Could not locate 'const pipetteData = [...];' in index.html (will be added).")
            
        # Replace pcData
        formatted_pc = format_json(pc_data)
        pattern_pc = r"(const pcData = )\[[\s\S]*?\];"
        replacement_pc = f"\\1{formatted_pc};"
        html_content, count_pc = re.subn(pattern_pc, replacement_pc, html_content)
        if count_pc == 0:
            print("Warning: Could not locate 'const pcData = [...];' in index.html (will be added).")

        formatted_animal_lab = format_json(animal_lab_data)
        pattern_animal_lab = r"(const animalLabData = )\[[\s\S]*?\];"
        replacement_animal_lab = f"\\1{formatted_animal_lab};"
        html_content, count_animal_lab = re.subn(pattern_animal_lab, replacement_animal_lab, html_content)
        if count_animal_lab == 0:
            print("Warning: Could not locate 'const animalLabData = [...];' in index.html.")

        formatted_reagent = format_json(reagent_data)
        pattern_reagent = r"(const reagentData = )\[[\s\S]*?\];"
        replacement_reagent = f"\\1{formatted_reagent};"
        html_content, count_reagent = re.subn(pattern_reagent, replacement_reagent, html_content)
        if count_reagent == 0:
            print("Warning: Could not locate 'const reagentData = [...];' in index.html.")
            
        meta_parts = [f"데이터 갱신: {datetime.date.today():%Y-%m-%d}"]
        equipment_stamp = re.search(r"\d{8}", excel_path.name)
        if equipment_stamp:
            meta_parts.append(f"장비목록 {equipment_stamp.group()}")
        reagent_stamp = re.search(r"v[\d.]+", reagent_files[-1].name) if reagent_files else None
        if reagent_stamp:
            meta_parts.append(f"시약목록 {reagent_stamp.group()}")
        meta_text = " · ".join(meta_parts)
        html_content = re.sub(
            r"<!--DATA_META-->.*?<!--/DATA_META-->",
            lambda _: f"<!--DATA_META-->{meta_text}<!--/DATA_META-->",
            html_content,
            flags=re.S,
        )

        html_path.write_text(html_content, encoding="utf-8")
        print("index.html has been successfully updated with new data.")
        return True
        
    except Exception as e:
        print(f"An error occurred while updating index.html: {e}")
        return False

if __name__ == "__main__":
    success = sync()
    if not success:
        sys.exit(1)
