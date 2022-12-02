from docx.elements.elements import TextElement
from docx.elements.paragraph import Paragraph
from docx.ooxml_ns import ns


class ParagraphGroup(TextElement):
    def __repr__(self):
        return f"ParagraphGroup(text='{self.text}')"

    def __iter__(self):
        return iter(self.paragraphs)

    def __str__(self):
        return self.text

    @property
    def text(self):
        return "\n".join(
            z
            for z in (
                "".join(run.text for run in para.runs) for para in self.paragraphs
            )
        )

    @property
    def paragraphs(self):
        return [Paragraph(el) for el in self.element.xpath("w:p", **ns)]


class Bubble(ParagraphGroup):
    pass
