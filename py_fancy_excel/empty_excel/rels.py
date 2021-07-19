import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element, tostring

class rel:
    def __init__(self, id:str="", type:str="", target:str="")->None:
        self.id:str = id
        self.type:str = type
        self.target:str = target
        self.tree:Element = self._get_tree("Relationship")

    def __str__(self)->str:
        return tostring(self.tree, "utf-8", method="xml", short_empty_elements=True).decode("utf-8")

    def from_str(self, xml_str:str)->None:
        tree:Element = ET.fromstring(xml_str)
        self.id:str = tree.attrib.get("Id") or ""
        self.type:str = tree.attrib.get("Type") or ""
        self.target:str = tree.attrib.get("Target") or ""
        self.tree:Element = tree

    def from_tree(self, tree:Element)->None:
        self.id:str = tree.attrib.get("Id") or ""
        self.type:str = tree.attrib.get("Type") or ""
        self.target:str = tree.attrib.get("Target") or ""
        self.tree:Element = tree

    def _get_tree(self, name:str)->Element:
        tree:Element = Element(name)
        tree.set("Id", self.id)
        tree.set("Type", self.type)
        tree.set("Target", self.target)
        return tree

    def get_tree(self)->Element:
        return self.tree


class rels:
    def __init__(self, rel_list:list=None, xmlns:str=""):
        self.xmlns = xmlns or "http://schemas.openxmlformats.org/package/2006/relationships"
        self.rel_list = rel_list if rel_list else []
        self.tree = self._get_tree("Relationships")

    def __str__(self):
        return tostring(self.tree, "utf-8", method="xml", short_empty_elements=True).decode("utf-8")

    def _get_tree(self, name:str)->Element:
        tree:Element = Element(name)
        tree.set("xmlns", self.xmlns)
        for rel in self.rel_list:
            tree.append(rel.get_tree())
        return tree

    def get_tree(self)->Element:
        return self.tree