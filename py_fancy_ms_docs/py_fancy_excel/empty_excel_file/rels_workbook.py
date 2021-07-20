from lxml.etree import Element, tostring
from .rels import rel, rels


r_id_1: rel = rel(
    "rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "worksheets/sheet1.xml")
r_id_2: rel = rel(
    "rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml")
r_id_3: rel = rel(
    "rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml")


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

    def __init__(self, version: str = "1.0", encoding: str = "UTF-8", standalone: str = "yes", key: str = "xl/_rels/workbook.xml.rels", rel_list: list = None) -> None:
        self.version: str = version
        self.encoding: str = encoding
        self.standalone: str = standalone
        self.key: str = key
        self.rel_list: list = rel_list or [r_id_3, r_id_2, r_id_1]

        self.rels: rels = rels(rel_list=self.rel_list)
        self.tree: Element = self.rels.get_tree()
        self.dict: dict = self._get_dict()

    def __str__(self) -> str:
        return f"{tostring(self.tree, encoding='UTF-8', xml_declaration=True, standalone=True).decode('utf-8')}"

    def _get_dict(self) -> dict:
        return {self.key: str(self)}
