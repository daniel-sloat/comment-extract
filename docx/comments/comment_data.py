"""Comment data containers"""

from ..elements import Paragraph, TextElement
from ..ooxml_ns import ns
from ..baseelement import BaseDOCXElement


class CommentMetaData(BaseDOCXElement):
    def __init__(self, element):
        super().__init__(element)
        self.bubble = Bubble(element)

    def __repr__(self):
        return f"CommentMetaData(tag='{self.tag}')"


class Bubble(TextElement):
    def __init__(self, element):
        self.element = element

    @property
    def paragraphs(self):
        return [Paragraph(el) for el in self.element.xpath("w:p", **ns)]


class CommentBounds:
    def __init__(self, start, end):
        self.start = start
        self.end = end

    def __repr__(self):
        return f"CommentBounds(_id='{self.start.xpath('string(@w:id)', **ns)}')"

    @property
    def _id(self):
        return self.start.xpath("string(@w:id)", **ns)

    @property
    def start_paragraph(self):
        return Paragraph(
            self.start.xpath("parent::w:p|following-sibling::w:p[1]", **ns)[0]
        )

    @property
    def end_paragraph(self):
        return Paragraph(
            self.end.xpath("parent::w:p|preceding-sibling::w:p[1]", **ns)[0]
        )

    @property
    def _para_id(self):
        return self.end_paragraph.attrib.get("paraId")

    # @property
    # def paragraphs(self):
    #     xpath = "(self::w:p|following-sibling::w:p)\
    #         [(w:r/w:t or w:commentRangeEnd) and not(re:test(string(.),'^\s+$'))]"
    #     paragraphs = (
    #         x
    #         for x in self.start_paragraph.node.xpath(
    #             xpath, namespaces=ooXMLns | regexns
    #         )
    #     )
    #     paras = []
    #     for para in paragraphs:
    #         paras.append(Paragraph(para))
    #         if para == self.end_paragraph.node:
    #             break
    #     return paras


class Comment:
    def __init__(self, start, end, id, bubble, author="", date="", initials=""):
        self.start = start
        self.end = end
        self._id = id
        self.author = author
        self.date = date
        self.initials = initials
        self.bubble = bubble

    def __repr__(self):
        return f"Comment(_id='{self.start.xpath('string(@w:id)', **ns)}')"

    @property
    def start_paragraph(self):
        return Paragraph(
            self.start.xpath("parent::w:p|following-sibling::w:p[1]", **ns)[0]
        )

    @property
    def end_paragraph(self):
        return Paragraph(
            self.end.xpath("parent::w:p|preceding-sibling::w:p[1]", **ns)[0]
        )

    @property
    def para_id(self):
        return self.end_paragraph.attrib.get("paraId")

    @property
    def paragraphs(self):
        xpath = "(self::w:p|following-sibling::w:p)\
            [(w:r/w:t or w:commentRangeEnd) and not(re:test(string(.),'^\s+$'))]"
        for para in self.start_paragraph.element.xpath(xpath, **ns):
            yield CommentParagraph(para, self._id)
            if para == self.end_paragraph.element:
                break


class CommentParagraph(Paragraph):
    def __init__(self, element, _id):
        super().__init__(element)
        self._id = _id

    @property
    def comment_runs(self):
        for run in (run for run in self.runs):
            if run.element.xpath("self::w:r", **ns):
                yield run
            elif run.element.xpath(f"self::w:commentRangeEnd[@w:id={self._id}]", **ns):
                break
