from ..baseelement import BaseDOCXElement
from ..ooxml_ns import ns


class StyleElement(BaseDOCXElement):
    def __init__(self, element):
        super().__init__(element)

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
        return self._name_tag(self.attrib.get("styleId"))

    @property
    def _name(self):
        return self.element.xpath("string(w:name/@w:val)", **ns)

    @property
    def _type(self):
        return self.attrib.get("type")

    @property
    def _basedon(self):
        return self.element.xpath("string(w:basedOn/@w:val)", **ns)
