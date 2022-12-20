"""Paragraph object"""
from lxml import etree

from docx.elements.attrib import get_attrib
from docx.elements.element_base import DOCXElement
from docx.elements.run import Run
from docx.ooxml_ns import ns


class Paragraph(DOCXElement):
    """Representation of <w:p> (paragraph) element."""

    def __init__(self, element, document):
        super().__init__(element)
        self._doc = document
        self.runs = [Run(run, self) for run in self.element.xpath("w:r", **ns)]

    def __getitem__(self, key):
        return self.runs[key]

    def __iter__(self):
        return iter(self.runs)

    def __len__(self):
        return len(self.runs)

    def __str__(self):
        return self.text

    def insert_run(self, position, element):
        self.runs.insert(position, Run(element, self._doc))

    @property
    def text(self):
        return "".join(run.text for run in self.runs)

    @property
    def props(self):
        return get_attrib(self.element.xpath("w:pPr/*", **ns))

    @property
    def glyph_props(self):
        return get_attrib(self.element.xpath("w:pPr/w:rPr/*", **ns))

    @property
    def style(self):
        return self.element.xpath("string(w:pPr/w:pStyle/@w:val)", **ns)

    @property
    def style_props(self):
        return self._doc.styles[self.style].paragraph


class CommentParagraph(Paragraph):
    def __init__(self, element, document, comment):
        super().__init__(element, document)
        self._comment = comment
        self.runs = self.get_comment_runs()

    def get_comment_runs(self):
        comment_run = etree.XPath(
            "self::w:r[w:t|w:footnoteReference|w:endnoteReference]", **ns
        )
        comment_start = etree.XPath(
            f"self::w:commentRangeStart[@w:id={self._comment._id}]", **ns
        )
        comment_end = etree.XPath(
            f"self::w:commentRangeEnd[@w:id={self._comment._id}]", **ns
        )

        def get_runs(children):
            runs = []
            for child in children:
                if comment_run(child):
                    runs.append(Run(child, self))
                elif comment_end(child):
                    break
            return runs

        para_elements = (el for el in self.element)
        comment_start_paragraph = self.element.xpath(
            f"w:commentRangeStart[@w:id={self._comment._id}]", **ns
        )

        if comment_start_paragraph:
            for element in para_elements:
                if comment_start(element):
                    return get_runs(para_elements)
        else:
            return get_runs(para_elements)

    # def get_comment_runs(self):
    #     comment_run = etree.XPath(
    #         "self::w:r[w:t|w:footnoteReference|w:endnoteReference]", **ns
    #     )
    #     comment_start = etree.XPath(
    #         f"self::w:commentRangeStart[@w:id={self._comment._id}]", **ns
    #     )
    #     comment_end = etree.XPath(
    #         f"self::w:commentRangeEnd[@w:id={self._comment._id}]", **ns
    #     )

    #     para_elements = (el for el in self.element)
    #     comment_start_paragraph = self.element.xpath(
    #         f"w:commentRangeStart[@w:id={self._comment._id}]", **ns
    #     )

    #     runs = []
    #     if comment_start_paragraph:
    #         for element in para_elements:
    #             if comment_start(element):
    #                 if comment_run(element):
    #                     runs.append(Run(element, self))
    #                 elif comment_end(element):
    #                     break
    #     else:
    #         for element in para_elements:
    #             if comment_run(element):
    #                 runs.append(Run(element, self))
    #             elif comment_end(element):
    #                 break
    #     return runs
