from docx.elements.elements import TextElement
from docx.elements.paragraph import Paragraph
from docx.ooxml_ns import ns
from reprlib import Repr

limit_repr_text = Repr()


class ParagraphGroup:
    def __init__(self):
        self.paragraphs = []

    def __repr__(self):
        return f"{self.__class__.__name__}(text='{limit_repr_text.repr(self.text)}')"
    
    def __str__(self):
        return self.text

    def __iter__(self):
        return iter(self.paragraphs)
    
    def __len__(self):
        return len(self.paragraphs)

    def __getitem__(self, key):
        return self.paragraphs[key]

    @property
    def text(self):
        return "\n".join(
            z
            for z in (
                "".join(run.text for run in para.runs) for para in self.paragraphs
            )
        )


class Bubble(ParagraphGroup):
    pass
