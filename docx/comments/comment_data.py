"""Comment data containers"""

from ..baseelement import BaseDOCXElement
from ..elements import CommentParagraph, Bubble
from ..ooxml_ns import ns


class CommentMetaData(BaseDOCXElement):
    def __init__(self, element):
        super().__init__(element)
        self.bubble = Bubble(element)

    @property
    def para_id(self):
        return self.element.xpath("string(w:p[last()]/@w14:paraId)", **ns)


class CommentBounds:
    def __init__(self, start, end):
        self.start = start
        self.end = end

    def __repr__(self):
        return f"CommentBounds(_id='{self._id}')"

    @property
    def _id(self):
        return self.start.xpath("string(@w:id)", **ns)

    @property
    def last_para_id(self):
        return self.end.xpath(
            "string((parent::w:p|preceding-sibling::w:p[1])/@w14:paraId)", **ns
        )


class Comment:
    def __init__(
        self, filename, start, end, _id, bubble, author="", date="", initials=""
    ):
        self.filename = filename
        self._start = start
        self._end = end
        self._id = _id
        self.author = author
        self.date = date
        self.initials = initials
        self.bubble = bubble

    def __repr__(self):
        return f"Comment(file='{self.filename}',_id='{self._id}')"

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
