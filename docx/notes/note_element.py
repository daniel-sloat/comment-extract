from docx.ooxml_ns import ns
from docx.elements import Paragraph, TextElement


class NoteElement(TextElement):
    def __repr__(self):
        return f"NoteElement(_id='{self._id}',_type='{self._type}')"

    @property
    def _id(self):
        return self.element.xpath("string(@w:id)", **ns)
    
    @property
    def _type(self):
        return self.element.xpath("string(@w:type)", **ns)

    @property
    def text(self):
        return "\n".join(z for z in ("".join(run.text for run in para.runs) for para in self.paragraphs))
    
    @property
    def paragraphs(self):
        return [Paragraph(el) for el in self.element]
