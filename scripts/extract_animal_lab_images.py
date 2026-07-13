"""Extract embedded photos from the animal-lab inventory workbook.

Excel stores the photos in this workbook as in-cell rich-data values rather
than ordinary worksheet images. This script resolves those values back to the
embedded media and saves one web-friendly JPEG per inventory row.
"""

from io import BytesIO
from pathlib import Path
from posixpath import normpath
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import openpyxl
from PIL import Image


WORKBOOK_NAME = "archive/동물실험실 관련 물품 리스트.xlsx"
SHEET_NAME = "Sheet1"
PHOTO_COLUMN = 2
DATA_START_ROW = 3

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
RICH_DATA_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
RICH_REL_NS = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"


def rich_value_images(package):
    """Return ``{value-metadata-index: image-bytes}`` for in-cell images."""
    rich_values = ET.fromstring(package.read("xl/richData/rdrichvalue.xml"))
    rich_relations = ET.fromstring(package.read("xl/richData/richValueRel.xml"))
    relation_file = ET.fromstring(package.read("xl/richData/_rels/richValueRel.xml.rels"))

    targets = {
        rel.attrib["Id"]: normpath(f"xl/richData/{rel.attrib['Target']}")
        for rel in relation_file.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
    }
    relation_ids = [
        rel.attrib[f"{{{DOC_REL_NS}}}id"]
        for rel in rich_relations.findall(f"{{{RICH_REL_NS}}}rel")
    ]

    images = {}
    for metadata_idx, rich_value in enumerate(
        rich_values.findall(f"{{{RICH_DATA_NS}}}rv"), start=1
    ):
        values = rich_value.findall(f"{{{RICH_DATA_NS}}}v")
        if not values:
            continue
        try:
            relation_idx = int(values[0].text)
            image_path = targets[relation_ids[relation_idx]]
            images[metadata_idx] = package.read(image_path)
        except (IndexError, KeyError, TypeError, ValueError):
            continue
    return images


def sheet_path(package):
    workbook = ET.fromstring(package.read("xl/workbook.xml"))
    relationships = ET.fromstring(package.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in relationships.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
    }
    for sheet in workbook.findall(f".//{{{MAIN_NS}}}sheet"):
        if sheet.attrib["name"].strip() == SHEET_NAME:
            relationship_id = sheet.attrib[f"{{{DOC_REL_NS}}}id"]
            return normpath(f"xl/{targets[relationship_id]}")
    raise ValueError(f"Could not find sheet named '{SHEET_NAME}'.")


def image_cells(package, worksheet_path):
    worksheet = ET.fromstring(package.read(worksheet_path))
    for cell in worksheet.findall(f".//{{{MAIN_NS}}}c"):
        metadata_idx = cell.attrib.get("vm")
        coordinate = cell.attrib.get("r", "")
        column = "".join(char for char in coordinate if char.isalpha())
        row = "".join(char for char in coordinate if char.isdigit())
        if not metadata_idx or not row:
            continue
        if openpyxl.utils.column_index_from_string(column) == PHOTO_COLUMN:
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
    if SHEET_NAME not in workbook.sheetnames:
        print(f"Error: Could not find sheet named '{SHEET_NAME}'.")
        workbook.close()
        return

    worksheet = workbook[SHEET_NAME]
    extracted_count = 0
    with ZipFile(xlsx_path) as package:
        try:
            images_by_metadata = rich_value_images(package)
            worksheet_path = sheet_path(package)
        except (KeyError, ET.ParseError, ValueError) as exc:
            print(f"Error: Could not read embedded image data: {exc}")
            workbook.close()
            return

        for row_idx, metadata_idx in image_cells(package, worksheet_path):
            if row_idx < DATA_START_ROW:
                continue
            image_bytes = images_by_metadata.get(metadata_idx)
            if image_bytes is None:
                print(f"Warning: No image data for row {row_idx}.")
                continue
            destination = images_dir / f"animal_lab_{row_idx - DATA_START_ROW + 1}.jpeg"
            try:
                save_as_jpeg(image_bytes, destination)
                print(f"Saved {destination.name} (row {row_idx})")
                extracted_count += 1
            except (OSError, ValueError) as exc:
                print(f"Error extracting image at row {row_idx}: {exc}")

    workbook.close()
    print(f"\nExtraction complete: saved {extracted_count} animal-lab images.")


if __name__ == "__main__":
    main()
