import os
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET


XML_HEADER = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n"""
WORKBOOK_HEADER = """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"""
WORKSHEET_HEADER = """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"""


class XLSXSheet:
    def __init__(self, _id, name, path):
        self.id = _id
        self.name = name
        self.path = path
        self.relationshipId = f"rId{_id}"

        self.file = open(path, "w", encoding='utf-8')
        self.file.write(XML_HEADER)
        self.file.write(WORKSHEET_HEADER)
        self.file.write("<sheetData>")

    def append_row(self, *columns):
        row = ET.Element("row")

        for column in columns:
            c = ET.SubElement(row, "c", {"t": "inlineStr"})
            s = ET.SubElement(c, "is")
            t = ET.SubElement(s, "t")
            t.text = column

        self.file.write(ET.tostring(row, encoding="unicode"))

    def finalize(self):
        self.file.write("</sheetData></worksheet>")
        self.file.close()


class XLSXBook:
    def __init__(self):
        self.base_dir = tempfile.mkdtemp()
        self.app_dir = os.path.join(self.base_dir, "xl")
        self.sheets = []

        os.mkdir(self.app_dir)
        os.mkdir(os.path.join(self.app_dir, "worksheets"))

    def add_sheet(self, name):
        _id = str(len(self.sheets) + 1)
        path = os.path.join(self.app_dir, f"worksheets/sheet{_id}.xml")
        sheet = XLSXSheet(_id, name, path)

        self.sheets.append(sheet)
        return sheet

    def _create_content_types(self):
        types = ET.Element("Types", {"xmlns": "http://schemas.openxmlformats.org/package/2006/content-types"})
        ET.SubElement(types, "Default", {"Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"})
        ET.SubElement(types, "Default", {"Extension": "xml", "ContentType": "application/xml"})
        ET.SubElement(types, "Override", {"PartName": "/xl/workbook.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"})

        for sheet in self.sheets:
            rel_path = sheet.path[len(self.base_dir):]
            ET.SubElement(types, "Override", {"PartName": rel_path, "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})

        with open(os.path.join(self.base_dir, "[Content_Types].xml"), 'w', encoding='utf-8') as f:
            f.write(XML_HEADER)
            f.write(ET.tostring(types, encoding="unicode"))

    def _create_root_rels(self):
        os.mkdir(os.path.join(self.base_dir, "_rels"))

        relationships = ET.Element("Relationships", {"xmlns": "http://schemas.openxmlformats.org/package/2006/relationships"})
        ET.SubElement(relationships, "Relationship", {
            "Id": "rId1",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "Target": "xl/workbook.xml"
        })

        with open(os.path.join(self.base_dir, "_rels/.rels"), 'w', encoding='utf-8') as f:
            f.write(XML_HEADER)
            f.write(ET.tostring(relationships, encoding="unicode"))

    def _create_app_rels(self):
        os.mkdir(os.path.join(self.app_dir, "_rels"))

        relationships = ET.Element("Relationships", {"xmlns": "http://schemas.openxmlformats.org/package/2006/relationships"})
        for sheet in self.sheets:
            rel_path = os.path.relpath(sheet.path, start=self.app_dir)

            ET.SubElement(relationships, "Relationship", {
                "Id": sheet.relationshipId,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                "Target": rel_path
            })

        with open(os.path.join(self.app_dir, "_rels/workbook.xml.rels"), 'w', encoding='utf-8') as f:
            f.write(XML_HEADER)
            f.write(ET.tostring(relationships, encoding="unicode"))

    def _create_workbook(self):
        sheets = ET.Element("sheets")
        for s, sheet in enumerate(self.sheets):
            ET.SubElement(sheets, "sheet", {"name": sheet.name, "sheetId": sheet.id, "r:id": sheet.relationshipId})

            with open(os.path.join(self.base_dir, "xl/workbook.xml"), 'w', encoding='utf-8') as f:
                f.write(XML_HEADER)
                f.write(WORKBOOK_HEADER)
                f.write(ET.tostring(sheets, encoding="unicode"))
                f.write("</workbook>")

    def _archive_dir(self, to_file):
        archive = zipfile.ZipFile(to_file, 'w', zipfile.ZIP_DEFLATED)

        for root, dirs, files in os.walk(self.base_dir):
            for file in files:
                rel_path = os.path.relpath(os.path.join(root, file), start=self.base_dir)
                archive.write(os.path.join(root, file), arcname=rel_path)

        archive.close()

    def finalize(self, to_file, remove_dir=True):
        # must have at least one sheet
        if not self.sheets:
            self.add_sheet("Sheet1")

        self._create_content_types()
        self._create_root_rels()
        self._create_app_rels()
        self._create_workbook()

        for sheet in self.sheets:
            sheet.finalize()

        self._archive_dir(to_file)

        if remove_dir:
            shutil.rmtree(self.base_dir)
