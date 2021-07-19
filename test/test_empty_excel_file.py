import unittest, sys   # The test framework
# from ..py_fancy_ms_docs import py_fancy_excel


class Test_EmptyExcelFile(unittest.TestCase):
    def setUp(self):
        sys.path.insert(0, "../py_fancy_ms_docs")
        from py_fancy_ms_docs import py_fancy_excel
        
    def test_str_method_relationships(self):
        test_str_relationships = str(py_fancy_excel.rels())
        original_empty_data = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>"
        self.assertEqual(test_str_relationships, original_empty_data)

    def test_str_method_relationships_workbook(self):
        test_str_relationships_workbook_dict  = py_fancy_excel.rels_workbook().get_dict()
        original_empty_data = {"xl/_rels/workbook.xml.rels": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>"}
        self.assertEqual(test_str_relationships_workbook_dict, original_empty_data)


# if __name__ == '__main__':
#     unittest.main()