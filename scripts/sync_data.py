#!/usr/bin/env python3
import os
import re
import sys
import json
from pathlib import Path
import openpyxl

def sync():
    root = Path(__file__).resolve().parents[1]
    excel_path = root / "담당자.xlsx"
    html_path = root / "index.html"
    
    if not excel_path.exists():
        print(f"Error: Excel file '{excel_path}' not found.")
        return False
        
    print(f"Reading data from {excel_path.name}...")
    try:
        wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
        sheet_name = "장비 목록 (전체)"
        if sheet_name not in wb.sheetnames:
            print(f"Error: Sheet '{sheet_name}' not found in Excel file.")
            print(f"Available sheets: {wb.sheetnames}")
            return False
            
        ws = wb[sheet_name]
        
        inventory_data = []
        # The data starts at row 4 (1-indexed)
        # Columns:
        # Col 2 (B): 번호 (no)
        # Col 3 (C): 장비명 (name)
        # Col 5 (E): 장비 위치 (loc)
        # Col 6 (F): 담당자 (mgr)
        # Col 7 (G): 비고 (remarks)
        
        for r_idx, row in enumerate(ws.iter_rows(min_row=4, values_only=True), start=4):
            if len(row) < 7:
                continue
                
            no_val = row[1]
            name_val = row[2]
            loc_val = row[4]
            mgr_val = row[5]
            remarks_val = row[6]
            
            # If no number is set, skip
            if no_val is None:
                continue
                
            try:
                no = int(no_val)
            except (ValueError, TypeError):
                # Skip header or non-numeric rows
                continue
                
            if not name_val:
                continue
                
            # Clean values
            loc = str(loc_val).strip() if loc_val is not None else ""
            if loc == "미지정":
                loc = ""
                
            mgr = str(mgr_val).strip() if mgr_val is not None else ""
            if mgr == "미지정":
                mgr = ""
                
            remarks = str(remarks_val).strip() if remarks_val is not None else ""
            
            inventory_data.append({
                "no": no,
                "name": str(name_val).strip(),
                "loc": loc,
                "mgr": mgr,
                "remarks": remarks
            })
            
        wb.close()
        
        print(f"Successfully parsed {len(inventory_data)} equipment entries.")
        
        # Write to index.html
        if not html_path.exists():
            print(f"Error: index.html not found at {html_path}")
            return False
            
        html_content = html_path.read_text(encoding="utf-8")
        
        # Format inventory data as JSON with exact indentation
        formatted_json = "[\n"
        for i, item in enumerate(inventory_data):
            comma = "," if i < len(inventory_data) - 1 else ""
            item_json = json.dumps(item, ensure_ascii=False)
            formatted_json += f"      {item_json}{comma}\n"
        formatted_json += "    ]"
        
        pattern = r"(const inventoryData = )\[[\s\S]*?\];"
        replacement = f"\\1{formatted_json};"
        
        updated_content, count = re.subn(pattern, replacement, html_content)
        if count == 0:
            print("Error: Could not locate 'const inventoryData = [...];' in index.html.")
            return False
            
        html_path.write_text(updated_content, encoding="utf-8")
        print("index.html has been successfully updated with new data.")
        return True
        
    except Exception as e:
        print(f"An error occurred: {e}")
        return False

if __name__ == "__main__":
    success = sync()
    if not success:
        sys.exit(1)
