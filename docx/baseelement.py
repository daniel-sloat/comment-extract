import re

from .ooxml_ns import *


class BaseDOCXElement:
    def __init__(self, element):
        self.element = element

    def __repr__(self):
        return f"DOCXElement(tag={self.tag})"

    @property
    def tag(self):
        return self._name_tag(self.element.tag)

    @property
    def attrib(self):
        return {self._name_tag(tag): el for tag, el in self.element.items()}

    # TODO This does not allow for multiple tags under the node (the dictionary
    # assures this). Need to refactor or use xpath.
    # @property
    # def child(self):
    #     return {
    #         self._name_tag(element.tag): BaseDOCXElement(element)
    #         for element in self.element
    #     }

    # @property
    # def parent(self):
    #     return BaseDOCXElement(self.element.getparent())

    @property
    def text(self):
        return self.element.xpath("string(.)", **ns)

    @staticmethod
    def _name_tag(tag):
        return re.sub("{.*}", "", tag if tag else "")
