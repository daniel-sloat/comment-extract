from .ooxml_ns import *
from .baseelement import BaseDOCXElement


class TextElement(BaseDOCXElement):
    def __init__(self, element):
        super().__init__(element)

    def __repr__(self):
        disp = 32
        return (
            f"{self.__class__.__name__}("
            f"text='{self.text[0:disp]}"
            f"{'...' if self.text and len(self.text) > disp else ''}"
            f"')"
        )


class Paragraph(TextElement):
    def __init__(self, element):
        super().__init__(element)

    @property
    def runs(self):
        return [Run(element) for element in self.element.xpath("w:r", **ns)]

    @property
    def props(self):
        return [PropElement(el) for el in self.element.xpath("w:pPr/*", **ns)]

    @property
    def style(self):
        return self.element.xpath("string(w:pPr/w:pStyle/@w:val)", **ns)


class Run(TextElement):
    def __init__(self, element):
        super().__init__(element)

    @property
    def props(self):
        return [PropElement(el) for el in self.element.xpath("w:rPr/*", **ns)]

    @property
    def style(self):
        return self.element.xpath("string(w:rPr/w:rStyle/@w:val)", **ns)


class PropElement(BaseDOCXElement):
    def __init__(self, element):
        self.element = element

    def __repr__(self):
        return f"PropElement(tag={self.tag})"
