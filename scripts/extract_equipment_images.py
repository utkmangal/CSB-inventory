"""Extract the in-cell equipment photos from ``담당자.xlsx``.

Recent Excel versions store pictures inserted with "Place in Cell" as rich-data
values.  Those images are not exposed through ``openpyxl.worksheet._images``,
so their order must be read from the XLSX package metadata instead.
"""

from io import BytesIO
from pathlib import Path
from posixpath import normpath
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import openpyxl
from PIL import Image


WORKBOOK_NAME = "담당자.xlsx"
TARGET_SHEET_NAME = "장비 목록 (전체)"
HEADER_ROW_LIMIT = 10

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
RICH_DATA_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
RICH_REL_NS = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"


def find_target_sheet(workbook):
    return next((ws for ws in workbook.worksheets if ws.title.strip() == TARGET_SHEET_NAME), None)


def find_header_row_and_columns(worksheet):
    for row_idx in range(1, min(HEADER_ROW_LIMIT, worksheet.max_row) + 1):
        headers = {str(cell.value).strip(): cell.column for cell in worksheet[row_idx] if cell.value is not None}
        if "번호" in headers and "사진" in headers:
            return row_idx, headers["번호"], headers["사진"]
    return None, None, None


def workbook_sheet_path(package, target_sheet_name):
    """Return the ZIP member for a sheet name, without assuming sheet numbers."""
    workbook = ET.fromstring(package.read("xl/workbook.xml"))
    relationships = ET.fromstring(package.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in relationships.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
    }
    for sheet in workbook.findall(f".//{{{MAIN_NS}}}sheet"):
        if sheet.attrib["name"].strip() == target_sheet_name:
            relationship_id = sheet.attrib[f"{{{DOC_REL_NS}}}id"]
            return normpath(f"xl/{targets[relationship_id]}")
    raise ValueError(f"Could not find sheet named '{target_sheet_name}'.")


def rich_value_images(package):
    """Return ``{value-metadata-index: image-bytes}`` for Excel in-cell images."""
    rich_values = ET.fromstring(package.read("xl/richData/rdrichvalue.xml"))
    rich_relations = ET.fromstring(package.read("xl/richData/richValueRel.xml"))
    relation_file = ET.fromstring(package.read("xl/richData/_rels/richValueRel.xml.rels"))

    targets = {
        rel.attrib["Id"]: normpath(f"xl/richData/{rel.attrib['Target']}")
        for rel in relation_file.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
    }
    relation_ids = [rel.attrib[f"{{{DOC_REL_NS}}}id"] for rel in rich_relations.findall(f"{{{RICH_REL_NS}}}rel")]

    images = {}
    for metadata_idx, rich_value in enumerate(rich_values.findall(f"{{{RICH_DATA_NS}}}rv"), start=1):
        values = rich_value.findall(f"{{{RICH_DATA_NS}}}v")
        if not values:
            continue
        try:
            relation_idx = int(values[0].text)
            image_path = targets[relation_ids[relation_idx]]
            images[metadata_idx] = package.read(image_path)
        except (IndexError, KeyError, TypeError, ValueError):
            # A rich value can be text or another non-image data type.
            continue
    return images


def image_cells(package, sheet_path, image_column_idx):
    """Yield (row, value-metadata-index) for rich images in the photo column."""
    worksheet = ET.fromstring(package.read(sheet_path))
    for cell in worksheet.findall(f".//{{{MAIN_NS}}}c"):
        metadata_idx = cell.attrib.get("vm")
        coordinate = cell.attrib.get("r", "")
        column = "".join(char for char in coordinate if char.isalpha())
        row = "".join(char for char in coordinate if char.isdigit())
        if not metadata_idx or not row or openpyxl.utils.column_index_from_string(column) != image_column_idx:
            continue
        yield int(row), int(metadata_idx)


def save_as_jpeg(image_bytes, destination):
    with Image.open(BytesIO(image_bytes)) as image:
        if image.mode != "RGB":
            image = image.convert("RGB")
        image.save(destination, "JPEG", quality=90)


def main():
    root = Path(__file__).resolve().parents[1]
    xlsx_path = root / WORKBOOK_NAME
    images_dir = root / "images"
    images_dir.mkdir(exist_ok=True)

    if not xlsx_path.exists():
        print(f"Error: {xlsx_path} not found.")
        return

    workbook = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    worksheet = find_target_sheet(workbook)
    if worksheet is None:
        print(f"Error: Could not find sheet named '{TARGET_SHEET_NAME}'.")
        workbook.close()
        return

    header_row_idx, number_col_idx, image_col_idx = find_header_row_and_columns(worksheet)
    if header_row_idx is None:
        print("Error: Could not locate the '번호' and '사진' headers.")
        workbook.close()
        return

    print(f"Loaded {xlsx_path.name}. Extracting in-cell images from '{worksheet.title}'...")
    extracted_count = 0
    with ZipFile(xlsx_path) as package:
        try:
            sheet_path = workbook_sheet_path(package, TARGET_SHEET_NAME)
            images_by_metadata = rich_value_images(package)
        except (KeyError, ET.ParseError, ValueError) as exc:
            print(f"Error: This workbook has no readable Excel in-cell image data: {exc}")
            workbook.close()
            return

        for row_idx, metadata_idx in image_cells(package, sheet_path, image_col_idx):
            equipment_no = worksheet.cell(row=row_idx, column=number_col_idx).value
            image_bytes = images_by_metadata.get(metadata_idx)
            if equipment_no is None or image_bytes is None:
                print(f"Warning: Skipped photo cell at row {row_idx} (missing number or image data).")
                continue
            try:
                equipment_no = int(equipment_no)
                destination = images_dir / f"equipment_{equipment_no}.jpeg"
                save_as_jpeg(image_bytes, destination)
                print(f"Saved {destination.name} (row {row_idx})")
                extracted_count += 1
            except (TypeError, ValueError, OSError) as exc:
                print(f"Error extracting image at row {row_idx}: {exc}")

    workbook.close()
    print(f"\nExtraction complete: saved {extracted_count} equipment images in equipment-number order.")


if __name__ == "__main__":
    main()
