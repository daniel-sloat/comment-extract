"""Styles for DOCX
    STYLE INHERITANCE
        PARAGRAPHS
            Use default paragraph properties (docDefaults)
            Append paragraph style properties
                [Local paragraph properties are only used for
                list formats and bullets, and are ignored.]

        RUNS
            Use default run properties (docDefaults)
            Append run style properties
            Append local run properties

        COMBINE PARAGRAPHS AND RUN FORMATTING
            Append result run properties over paragraph properties

    Styles can also be based on other styles, and 'inherit' those
    styles' format attributes. And that inherited style may itself
    be based on another style - and so on until the 'base style'.
"""

from ..ooxml_ns import ns
from .style_element import StyleElement
from ..elements import PropElement


class Styles:
    def __init__(self, document):
        self._doc = document
        self._style_xml = self._doc.xml["word/styles.xml"]

    def __repr__(self):
        return f"Styles(file='{self._doc.file}',count={len(self.styles)})"

    def __getitem__(self, key):
        return self.styles[key]

    def __iter__(self):
        return iter(self.styles.values())

    @property
    def styles(self):
        return {
            element.xpath("string(@w:styleId)", **ns): StyleElement(element)
            for element in self._style_xml.xpath("w:style", **ns)
        }

    @property
    def doc_default_props_para(self):
        return [
            PropElement(el)
            for el in self._style_xml.xpath("w:docDefaults/w:pPrDefault/w:pPr/*", **ns)
        ]

    @property
    def doc_default_props_run(self):
        return [
            PropElement(el)
            for el in self._style_xml.xpath("w:docDefaults/w:rPrDefault/w:rPr/*", **ns)
        ]
