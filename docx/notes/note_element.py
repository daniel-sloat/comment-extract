from docx.ooxml_ns import ns
from docx.elements.paragraph_group import ParagraphGroup
from docx.elements.paragraph import ParagraphStyled
class NoteElement(ParagraphGroup):
    def __init__(self, element, notes):
        super().__init__(element)
        self._styles = notes._doc.styles
        self._notes = notes

    def __repr__(self):
        return f"NoteElement(_id='{self._id}',_type='{self._type}')"

    @property
    def _id(self):
        return self.element.xpath("string(@w:id)", **ns)

    @property
    def _type(self):
        return self.element.xpath("string(@w:type)", **ns)

    # @property
    # def paragraphs(self):
    #     return [
    #         ParagraphStyled(el, self._styles) for el in self.element.xpath("w:p", **ns)
    #     ]
