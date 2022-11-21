from docx.baseelement import BaseDOCXElement
from docx.elements import PropElement
from docx.ooxml_ns import ns


class StyleElement(BaseDOCXElement):
    def __init__(self, element):
        super().__init__(element)
        self._combined_para = {}
        self._combined_run = {}
        self.basedon = self.element.xpath("string(w:basedOn/@w:val)", **ns)

    def __repr__(self):
        return (
            f"Style("
            f"id='{self._id}'"
            f",name='{self._name}'"
            f",type='{self._type}'"
            # f",basedon='{self._basedon}'"
            f")"
        )

    @property
    def _id(self):
        return self._strip_uri(self.attrib.get("styleId"))

    @property
    def _name(self):
        return self.element.xpath("string(w:name/@w:val)", **ns)

    @property
    def _type(self):
        return self.attrib.get("type")

    @property
    def basedon(self):
        return self._basedon

    @property
    def _paragraph(self):
        return {
            PropElement(el).tag: PropElement(el).attrib
            for el in self.element.xpath("w:pPr/*", **ns)
        }

    @property
    def _run(self):
        return {
            PropElement(el).tag: PropElement(el).attrib
            for el in self.element.xpath("w:rPr/*", **ns)
        }

    @property
    def _para_and_run(self):
        return {"para": self._paragraph, "run": self._run}

    @_paragraph.setter
    def _paragraph(self, d):
        self._paragraph = d

    @basedon.setter
    def basedon(self, b):
        self._basedon = b

