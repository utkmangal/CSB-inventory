"""Extract equipment photos from the latest equipment-list workbook.

The workbook (``*장비목록*.xlsx``) has the columns ``No | 장비명 | 설치 위치``
and the photos are floating pictures anchored next to each row.  Each picture
is mapped to the equipment number of the row it is anchored on.
"""

from io import BytesIO
from pathlib import Path
from posixpath import normpath
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import openpyxl
from PIL import Image

WORKBOOK_GLOB = "*장비목록*.xlsx"
ROW_HEIGHT_EMU = 1008000  # rows are 80.1pt high, pictures are 1008000 EMU tall

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DRAWING_MAIN_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def find_workbook(root):
    candidates = sorted(
        (p for p in root.glob(WORKBOOK_GLOB) if p.is_file() and not p.name.startswith("~$")),
        key=lambda p: p.name,
    )
    return candidates[-1] if candidates else None


def relationships(package, rels_path, base_dir):
    root = ET.fromstring(package.read(rels_path))
    return {
        rel.attrib["Id"]: normpath(f"{base_dir}/{rel.attrib['Target']}")
        for rel in root.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
    }


def first_sheet_path(package):
    workbook = ET.fromstring(package.read("xl/workbook.xml"))
    targets = relationships(package, "xl/_rels/workbook.xml.rels", "xl")
    sheet = workbook.find(f".//{{{MAIN_NS}}}sheet")
    return targets[sheet.attrib[f"{{{DOC_REL_NS}}}id"]]


def sheet_drawing_path(package, sheet_path):
    sheet = ET.fromstring(package.read(sheet_path))
    drawing = sheet.find(f"{{{MAIN_NS}}}drawing")
    if drawing is None:
        return None
    sheet_dir, sheet_file = sheet_path.rsplit("/", 1)
    targets = relationships(package, f"{sheet_dir}/_rels/{sheet_file}.rels", sheet_dir)
    return targets[drawing.attrib[f"{{{DOC_REL_NS}}}id"]]


def anchored_pictures(package, drawing_path):
    """Yield (0-based row index, image bytes) for every picture in the drawing."""
    drawing_dir, drawing_file = drawing_path.rsplit("/", 1)
    targets = relationships(package, f"{drawing_dir}/_rels/{drawing_file}.rels", drawing_dir)
    drawing = ET.fromstring(package.read(drawing_path))
    for anchor in drawing:
        start = anchor.find(f"{{{DRAWING_NS}}}from")
        blip = anchor.find(f".//{{{DRAWING_MAIN_NS}}}blip")
        if start is None or blip is None:
            continue
        image_path = targets.get(blip.attrib.get(f"{{{DOC_REL_NS}}}embed"))
        if image_path is None:
            continue
        row = int(start.find(f"{{{DRAWING_NS}}}row").text)
        row_offset = int(start.find(f"{{{DRAWING_NS}}}rowOff").text)
        # A picture shifted down by a full row height belongs to the next row.
        yield row + round(row_offset / ROW_HEIGHT_EMU), package.read(image_path)


def save_as_jpeg(image_bytes, destination):
    with Image.open(BytesIO(image_bytes)) as image:
        if image.mode != "RGB":
            image = image.convert("RGB")
        image.thumbnail((1600, 1600), Image.LANCZOS)
        image.save(destination, "JPEG", quality=82, optimize=True, progressive=True)


def main():
    root = Path(__file__).resolve().parents[1]
    xlsx_path = find_workbook(root)
    images_dir = root / "images"
    images_dir.mkdir(exist_ok=True)

    if xlsx_path is None:
        print(f"Error: no workbook matching '{WORKBOOK_GLOB}' found in {root}.")
        return 1

    workbook = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    numbers = {
        row_idx: row[0]
        for row_idx, row in enumerate(workbook.worksheets[0].iter_rows(max_col=1, values_only=True), start=1)
    }
    workbook.close()

    print(f"Loaded {xlsx_path.name}. Extracting equipment photos...")
    extracted_count = 0
    with ZipFile(xlsx_path) as package:
        drawing_path = sheet_drawing_path(package, first_sheet_path(package))
        if drawing_path is None:
            print("Error: the workbook sheet contains no pictures.")
            return 1

        for zero_based_row, image_bytes in anchored_pictures(package, drawing_path):
            row_idx = zero_based_row + 1
            equipment_no = numbers.get(row_idx)
            try:
                equipment_no = int(equipment_no)
            except (TypeError, ValueError):
                print(f"Warning: Skipped picture at row {row_idx} (no equipment number).")
                continue
            try:
                destination = images_dir / f"equipment_{equipment_no}.jpeg"
                save_as_jpeg(image_bytes, destination)
                for stale in (images_dir / f"equipment_{equipment_no}.png", images_dir / f"equipment_{equipment_no}.jpg"):
                    stale.unlink(missing_ok=True)
                print(f"Saved {destination.name} (row {row_idx})")
                extracted_count += 1
            except OSError as exc:
                print(f"Error extracting image at row {row_idx}: {exc}")

    print(f"\nExtraction complete: saved {extracted_count} equipment images.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
