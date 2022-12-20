"""Styles for DOCX
    STYLE INHERITANCE
        PARAGRAPHS
            Use default paragraph properties (docDefaults)
            Append paragraph style properties
                [Local paragraph properties are only used for list formats and bullets]

        RUNS
            Use default run properties (docDefaults)
            Append run style properties
            Append local run properties

        COMBINE PARAGRAPHS AND RUN FORMATTING
            Append result run properties over paragraph properties

    Styles can also be based on other styles, and 'inherit' those styles' format
    attributes. And that inherited style may itself be based on another style - and so
    on until the 'base style'.
"""

from docx.elements.attrib import get_attrib
from docx.ooxml_ns import ns
from docx.styles.style import Style


class Styles:
    """Represents styles data in OOXML document."""

    def __init__(self, document):
        self._doc = document
        self._style_xml = self._doc.xml["word/styles.xml"]
        self.style_ids = self._style_xml.xpath("w:style/@w:styleId", **ns)
        self.styles = {_id: Style(_id, self) for _id in self.style_ids}

    def __repr__(self):
        return f"Styles(file='{self._doc.file}',count={len(self)})"

    def __getitem__(self, key):
        return self.styles[key]

    def __iter__(self):
        return iter(self.styles.items())

    def __len__(self):
        return len(self.styles)

    @property
    def doc_default_props_para(self):
        xpath = "w:docDefaults/w:pPrDefault/w:pPr/*"
        return get_attrib(self._style_xml.xpath(xpath, **ns))

    @property
    def doc_default_props_run(self):
        xpath = "w:docDefaults/w:rPrDefault/w:rPr/*"
        return get_attrib(self._style_xml.xpath(xpath, **ns))
