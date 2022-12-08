from lxml.etree import XPath

from docx.elements.elements import DOCXElement
from docx.elements.run import Run
from docx.ooxml_ns import ns

from lxml.etree import QName


class Paragraph(DOCXElement):
    def __init__(self, element, group):
        super().__init__(element)
        self._group = group
        self.runs = [Run(run, self) for run in self.get_runs()]

    def __getitem__(self, key):
        return self.runs[key]

    def __iter__(self):
        return iter(self.runs)

    @property
    def text(self):
        return "".join(run.text for run in self.runs)

    def get_runs(self):
        return self.element.xpath("w:r", **ns)

    @property
    def props(self):
        return {
            QName(el).localname: el.attrib for el in self.element.xpath("w:pPr/*", **ns)
        }

    @property
    def glyph_props(self):
        return {
            QName(el).localname: el.attrib
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

    @property
    def style_props(self):
        styles = self._group._comments._doc.styles
        return styles.styles_map.get(self.style, {}).get("para", {})


class CommentParagraph(Paragraph):
    def __init__(self, element, comment):
        super().__init__(element, comment)
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
                    runs.append(Run(child, self))
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
