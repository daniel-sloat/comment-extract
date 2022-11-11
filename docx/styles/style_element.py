from ..baseelement import BaseDOCXElement


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
        return self.child.get("name").attrib.get("val")

    @property
    def _type(self):
        return self.attrib.get("type")

    @property
    def _basedon(self):
        return el.attrib.get("val") if (el := self.child.get("basedOn")) else ""
