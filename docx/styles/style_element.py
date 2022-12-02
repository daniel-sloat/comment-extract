from docx.elements.elements import BaseDOCXElement, PropElement
from docx.ooxml_ns import ns


class StyleElement(BaseDOCXElement):
    def __init__(self, element):
        super().__init__(element)
        self.basedon = self.element.xpath("string(w:basedOn/@w:val)", **ns)

    def __repr__(self):
        return (
            f"Style("
            f"id='{self._id}'"
            f",name='{self._name}'"
            f",type='{self._type}'"
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

    @basedon.setter
    def basedon(self, b):
        self._basedon = b

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
