from docx.elements.paragraph import Paragraph
from docx.elements.paragraph_group import ParagraphGroup
from docx.ooxml_ns import ns


class Note(ParagraphGroup):
    def __init__(self, _id, notes):
        self._id = _id
        self._parent = notes

    def __repr__(self):
        return f"{self.__class__.__name__}(_id='{self._id}')"


class FootNote(Note):
    def __init__(self, _id, notes):
        super().__init__(_id, notes)
        self.element = self._parent._footnotes_xml.xpath(
            "w:footnote[@w:id=$_id]", _id=self._id, **ns
        )[0]
        self.paragraphs = [
            Paragraph(para, self._parent._doc)
            for para in self.element.xpath("w:p", **ns)
        ]


class EndNote(Note):
    def __init__(self, _id, notes):
        super().__init__(_id, notes)
        self.element = self._parent._endnotes_xml.xpath(
            "w:endnote[@w:id=$_id]", _id=self._id, **ns
        )[0]
        self.paragraphs = [
            Paragraph(para, self._parent._doc)
            for para in self.element.xpath("w:p", **ns)
        ]
