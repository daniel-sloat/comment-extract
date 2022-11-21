"""Comment data containers"""

from lxml import etree
from docx.baseelement import BaseDOCXElement
from docx.elements import CommentParagraph, Bubble
from docx.ooxml_ns import ns


class CommentMetaData(BaseDOCXElement):
    def __init__(self, element):
        super().__init__(element)
        self.bubble = Bubble(element)

    @property
    def para_id(self):
        return self.element.xpath("string(w:p[last()]/@w14:paraId)", **ns)

    def asdict(self):
        return {
            **self.attrib,
            "bubble": self.bubble,
            "para_id": self.para_id,
        }


class CommentBounds:
    def __init__(self, start, end):
        self.start: etree._Element = start
        self.end: etree._Element = end

    @property
    def _id(self):
        return self.start.xpath("string(@w:id)", **ns)

    def asdict(self):
        return {
            "start": self.start,
            "end": self.end,
        }


class Comment:
    def __init__(
        self, start, end, _id, bubble, author="", date="", initials="", **kwargs
    ):
        self._start = start
        self._end = end
        self._id = _id
        self.author = author
        self.date = date
        self.initials = initials
        self.bubble = bubble

    def __repr__(self):
        return f"Comment(_id='{self._id}',text='{self.text}')"

    def __iter__(self):
        return iter(self.paragraphs)

    @property
    def text(self):
        return "".join(run.text for para in self.paragraphs for run in para.runs)

    @property
    def paragraphs(self):
        start_paragraph = self._start.xpath(
            "parent::w:p|following-sibling::w:p[1]",
            **ns,
        )[0]
        end_paragraph = self._end.xpath(
            "parent::w:p|preceding-sibling::w:p[1]",
            **ns,
        )[0]
        xpath = (
            "(self::w:p|following-sibling::w:p)"
            "[(w:r/w:t or w:commentRangeEnd) and not(re:test(string(.),'^\s+$'))]"
        )
        paragraphs = (x for x in start_paragraph.xpath(xpath, **ns))
        paras = []
        for para in paragraphs:
            paras.append(CommentParagraph(para, self._id))
            if para == end_paragraph:
                break
        return paras
