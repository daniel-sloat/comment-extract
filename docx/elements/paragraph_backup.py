from lxml.etree import XPath

from docx.elements.elements import PropElement, TextElement
from docx.elements.run import CommentRun, Run, RunStyled
from docx.ooxml_ns import ns


class Paragraph(TextElement):
    def __getitem__(self, key):
        return self.runs[key]

    def __iter__(self):
        return iter(self.runs)

    @property
    def text(self):
        return "".join(run.text for run in self.runs)

    @property
    def runs(self):
        return [Run(element) for element in self.element.xpath("w:r", **ns)]

    @property
    def props(self):
        return {
            (element := PropElement(el)).tag: element.attrib
            for el in self.element.xpath("w:pPr/*", **ns)
            if not el.xpath("self::w:rPr", **ns)
        }

    @property
    def glyph_props(self):
        return {
            (element := PropElement(el)).tag: element.attrib
            for el in self.element.xpath("w:pPr/w:rPr/*", **ns)
        }

    @property
    def style(self):
        return self.element.xpath("string(w:pPr/w:pStyle/@w:val)", **ns)

    @property
    def footnotes(self):
        return [run.footnote for run in self.runs if run.footnote]

    @property
    def endnotes(self):
        return [run.endnote for run in self.runs if run.endnote]


class ParagraphStyled(Paragraph):
    def __init__(self, element, styles):
        super().__init__(element)
        self._styles = styles

    @property
    def style_props(self):
        return self._styles.styles_map.get(self.style, {}).get("para", {})

    @property
    def runs(self):
        return [
            RunStyled(element, self._styles)
            for element in self.element.xpath("w:r", **ns)
        ]


class CommentParagraph(ParagraphStyled):
    def __init__(self, element, comment):
        super().__init__(element, comment._comments._doc.styles)
        self._comment = comment

    @property
    def runs(self):
        comment_run = XPath(
            "self::w:r[w:t|w:footnoteReference|w:endnoteReference]", **ns
        )
        comment_start = XPath(
            f"self::w:commentRangeStart[@w:id={self._comment._id}]", **ns
        )
        comment_end = XPath(f"self::w:commentRangeEnd[@w:id={self._comment._id}]", **ns)

        def get_runs(children):
            runs = []
            for child in children:
                if comment_run(child):
                    runs.append(CommentRun(child, self))
                elif comment_end(child):
                    break
            return runs

        para_elements = (
            el
            for el in self.element.xpath(
                f"w:r|w:commentRangeStart[@w:id={self._comment._id}]|w:commentRangeEnd[@w:id={self._comment._id}]",
                **ns,
            )
        )
        comment_start_paragraph = self.element.xpath(
            f"w:commentRangeStart[@w:id={self._comment._id}]", **ns
        )

        if comment_start_paragraph:
            for element in para_elements:
                if comment_start(element):
                    return get_runs(para_elements)
        else:
            return get_runs(para_elements)
