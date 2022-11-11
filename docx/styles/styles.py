from ..ooxml_ns import *
from .style_element import StyleElement


class Styles:
    def __init__(self, document):
        self._doc = document

    def __repr__(self):
        return self.styles

    @property
    def styles(self):
        style_elements = self._doc.xml["word/styles.xml"].iterfind("w:style", **ns)
        return {
            element.xpath("string(@w:styleId)", **ns): StyleElement(element)
            for element in style_elements
        }
