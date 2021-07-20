import sys
import unittest
import py_fancy_ms_docs
from lxml.etree import tostring, fromstring
sys.path.insert(0, "../py_fancy_ms_docs")
DEBUG = False


class test_empty_excel_file(unittest.TestCase):
    def test_str_method_relationship(self):
        """ Test if the new generated Relationship data matches the original """
        test_str_relationships = str(py_fancy_ms_docs.py_fancy_excel.rel(
            "rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "worksheets/sheet1.xml"))
        original_empty_data = "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"

        test_str_relationships = self._format_str(test_str_relationships)
        original_empty_data = self._format_str(original_empty_data)

        self.assertEqual(test_str_relationships, original_empty_data)

    def test_str_method_relationships(self):
        """ Test if the new generated Relationships data matches the original """
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

        test_str_relationships = self._format_str(test_str_relationships)
        original_empty_data = self._format_str(original_empty_data)

        self.assertEqual(test_str_relationships, original_empty_data)

    def test_str_method_relationships_workbook(self):
        """ Test if the new generated Relationships Workbook data matches the original """
        if DEBUG:
            self.maxDiff = None
        test_str_relationships_workbook_dict = py_fancy_ms_docs.py_fancy_excel.rels_workbook().dict
        original_empty_data = {"xl/_rels/workbook.xml.rels": "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>"}

        test_str_relationships_workbook_dict = self._format_dict(
            test_str_relationships_workbook_dict)
        original_empty_data = self._format_dict(original_empty_data)

        self.assertEqual(test_str_relationships_workbook_dict,
                         original_empty_data)

    def _format_str(self, _str):
        """ Function to remove an fromat from the string like different quotations """
        return tostring(fromstring(_str.encode("UTF-8"))).decode("UTF-8")

    def _format_dict(self, _dict):
        """ Function to remove an fromat from the dikt like rels_workbook() """
        return {k: self._format_str(v) for k, v in _dict.items()}
