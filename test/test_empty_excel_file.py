import sys
import unittest
import py_fancy_ms_docs
sys.path.insert(0, "../py_fancy_ms_docs")
DEBUG = False


class test_empty_excel_file(unittest.TestCase):
    def test_str_method_relationship(self):
        test_str_relationships = str(py_fancy_ms_docs.py_fancy_excel.rel(
            "rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "worksheets/sheet1.xml"))
        original_empty_data = "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
        self.assertEqual(test_str_relationships, original_empty_data)

    def test_str_method_relationships(self):
        if DEBUG:
            self.maxDiff = None
        rel_list = [
            py_fancy_ms_docs.py_fancy_excel.rel(
                "rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml"),
            py_fancy_ms_docs.py_fancy_excel.rel(
                "rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml"),
            py_fancy_ms_docs.py_fancy_excel.rel(
                "rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "worksheets/sheet1.xml")
        ]
        test_str_relationships = str(
            py_fancy_ms_docs.py_fancy_excel.rels(rel_list=rel_list))
        original_empty_data = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>"
        self.assertEqual(test_str_relationships, original_empty_data)

    def test_str_method_relationships_workbook(self):
        if DEBUG:
            self.maxDiff = None
        test_str_relationships_workbook_dict = py_fancy_ms_docs.py_fancy_excel.rels_workbook().dict
        original_empty_data = {"xl/_rels/workbook.xml.rels": "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>"}
        self.assertEqual(test_str_relationships_workbook_dict,
                         original_empty_data)
