# import xml.etree.ElementTree as ET
# from xml.etree.ElementTree import Element, tostring
from lxml.etree import Element, tostring


class rel:
    """ Excel Relationship Element """

    def __init__(self, id: str = "", type: str = "", target: str = "") -> None:
        self.id: str = id
        self.type: str = type
        self.target: str = target
        self.tree: Element = self._get_tree()

    def __str__(self) -> str:
        """ Get the Element string of Relationships Element """
        return tostring(self.tree, encoding="UTF-8").decode("UTF-8")

    def from_str(self, xml_str: str) -> None:
        """ Parse a string to Relationship Element """
        tree: Element = ET.fromstring(xml_str)
        self.id: str = tree.attrib.get("Id") or ""
        self.type: str = tree.attrib.get("Type") or ""
        self.target: str = tree.attrib.get("Target") or ""
        self.tree: Element = tree

    def from_tree(self, tree: Element) -> None:
        """ Parse a Element to Relationship Element """
        self.id: str = tree.attrib.get("Id") or ""
        self.type: str = tree.attrib.get("Type") or ""
        self.target: str = tree.attrib.get("Target") or ""
        self.tree: Element = tree

    def _get_tree(self) -> Element:
        """ Private method to get the Element tree of Relationship Element from self """
        tree: Element = Element("Relationship")
        tree.set("Id", self.id)
        tree.set("Type", self.type)
        tree.set("Target", self.target)
        return tree

    def get_tree(self) -> Element:
        """ Get the Element tree of Relationship Element """
        return self.tree


class rels:
    """ Excel Relationships Element """

    def __init__(self, rel_list: list = None, xmlns: str = "") -> None:
        self.xmlns = xmlns or "http://schemas.openxmlformats.org/package/2006/relationships"
        self.rel_list = rel_list if rel_list else []
        self.tree = self._get_tree()

    def __str__(self) -> str:
        """ Get the Element string of Relationships Element """
        return tostring(self.tree, encoding="UTF-8").decode("UTF-8")

    def _get_tree(self) -> Element:
        """ Private method to get the Element tree of Relationships Element from self """
        tree: Element = Element("Relationships")
        tree.set("xmlns", self.xmlns)
        for rel in self.rel_list:
            tree.append(rel.get_tree())
        return tree

    def get_tree(self) -> Element:
        """ Get the Element tree of Relationships Element """
        return self.tree
