from docx.ooxml_ns import ns
from docx.baseelement import BaseDOCXElement


class TextElement(BaseDOCXElement):
    def __repr__(self):
        disp = 32
        return (
            f"{self.__class__.__name__}("
            f"text='{self.text[0:disp]}"
            f"{'...' if self.text and len(self.text) > disp else ''}"
            f"')"
        )

    @property
    def text(self):
        return self.element.xpath("string(.)")


class Bubble(TextElement):
    @property
    def paragraphs(self):
        return [Paragraph(el) for el in self.element.xpath("w:p", **ns)]


class Paragraph(TextElement):
    def __iter__(self):
        return iter(self.runs)

    @property
    def runs(self):
        return [Run(element) for element in self.element.xpath("w:r", **ns)]

    @property
    def props(self):
        return [PropElement(el) for el in self.element.xpath("w:pPr/*", **ns)]

    @property
    def style(self):
        return self.element.xpath("string(w:pPr/w:pStyle/@w:val)", **ns)


class CommentParagraph(Paragraph):
    def __init__(self, element, _id):
        super().__init__(element)
        self._id = _id

    @property
    def text(self):
        return "".join(run.text for run in self.runs)

    @property
    def runs(self):
        def get_runs(children):
            select_runs = []
            for child in children:
                if child.xpath("*", **ns):
                    select_runs.append(Run(child))
                elif child.xpath(f"self::w:commentRangeEnd[@w:id={self._id}]", **ns):
                    break
            return select_runs

        para_elements = (el for el in self.element.xpath("*", **ns))
        comment_start_paragraph = self.element.xpath(
            f"boolean(w:commentRangeStart[@w:id={self._id}])", **ns
        )

        if comment_start_paragraph:
            for element in para_elements:
                if element.xpath(
                    f"boolean(self::w:commentRangeStart[@w:id={self._id}])", **ns
                ):
                    return get_runs(para_elements)
        else:
            return get_runs(para_elements)


class Run(TextElement):
    @property
    def props(self):
        return [PropElement(el) for el in self.element.xpath("w:rPr/*", **ns)]

    @property
    def style(self):
        return self.element.xpath("string(w:rPr/w:rStyle/@w:val)", **ns)


class PropElement(BaseDOCXElement):
    pass
