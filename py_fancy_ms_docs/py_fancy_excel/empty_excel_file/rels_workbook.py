from .rels import rel, rels


class rels_workbook:
    """
    Representing "xl/_rels/workbook.xml.rels" in a empty excel file.
    
    Reslt of self.get_dict(), but the result will not be formatted like this:
        {"xl/_rels/workbook.xml.rels": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n
            <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
                <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>
                <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>
                <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>
            </Relationships>"},
    """

    def __init__(self, version: str = "1.0", encoding: str = "UTF-8", standalone: str = "yes", key: str = "xl/_rels/workbook.xml.rels", rel_list: list = None):
        self.version = version
        self.encoding = encoding
        self.standalone = standalone
        self.key = key
        self.rel_list = rel_list or [
            rel("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml"),
            rel("rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml"),
            rel("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "worksheets/sheet1.xml")
        ]

        self.relationships = rels(rel_list=rel_list)
        self.dict = self._get_dict()

    def _get_str(self):
        str_rels_workbook = f"<?xml version=\"{self.version}\" encoding=\"{self.encoding}\" standalone=\"{self.standalone}\"?>\r\n"
        return "".join([str_rels_workbook, str(self.relationships)])

    def _get_dict(self):
        return {self.key: self._get_str()}