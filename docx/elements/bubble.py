from docx.elements.paragraph import Paragraph
from docx.elements.paragraph_group import ParagraphGroup


class Bubble(ParagraphGroup):
    def __init__(self, paragraphs, document):
        self._doc = document
        self.paragraphs = [Paragraph(el, self._doc) for el in paragraphs]
        super().__init__(self.paragraphs)
