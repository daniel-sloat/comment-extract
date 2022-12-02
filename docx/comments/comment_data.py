"""Comment data containers"""

from functools import cached_property
from lxml import etree
from docx.elements.elements import BaseDOCXElement
from docx.elements.paragraph import CommentParagraph
from docx.elements.paragraph_group import Bubble, ParagraphGroup
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


class Comment(ParagraphGroup):
    def __init__(
        self,
        *,
        start,
        end,
        _id,
        bubble,
        author="",
        date="",
        initials="",
        comments,
        **attrs,
    ):
        self._start = start
        self._end = end
        self._id = _id
        self.author = author
        self.date = date
        self.initials = initials
        self.bubble = bubble
        self._attrs = attrs
        self._comments = comments

    def __repr__(self):
        return f"Comment(_id='{self._id}')"
    
    def __getitem__(self, key):
        return self.paragraphs[key]

    @cached_property
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
            "[(not(re:test(string(.),'^\s*$')) or w:commentRangeEnd)]"
        )
        paragraphs = (x for x in start_paragraph.xpath(xpath, _id=self._id, **ns))
        paras = []
        for para in paragraphs:
            paras.append(CommentParagraph(para, self))
            if para == end_paragraph:
                break
        return paras

    # @cached_property
    # def paragraphs(self):
    #     end_paragraph2 = self._start.xpath(
    #         "(//w:p[w:commentRangeEnd/@w:id=$_id]|//w:p[following-sibling::w:commentRangeEnd/@w:id=$_id])",
    #         _id=self._id,
    #         **ns,
    #     )[0]
    #     # print(end_paragraph.xpath("string(.)", **ns))
    #     end_paragraph = self._end.xpath(
    #         "parent::w:p|preceding-sibling::w:p[1]",
    #         **ns,
    #     )[0]
    #     if end_paragraph != end_paragraph2:
    #         print(end_paragraph2.xpath("string(preceding-sibling::w:p[1])",**ns), self.bubble)
    #     xpath = (
    #         "//w:p[w:commentRangeStart/@w:id=$_id]/self::w:p"# and //w:p[preceding-sibling::w:commentRangeStart/@w:id=$_id]/following-siblings::w:p)"
    #         # "[(not(re:test(string(.),'^\s*$')) or w:commentRangeEnd)]"
    #     )
    #     paragraphs = (x for x in self._start.xpath(xpath, _id=self._id, **ns))
    #     paras = []
    #     for para in paragraphs:
    #         paras.append(CommentParagraph(para, self))
    #         if para == end_paragraph:
    #             break
    #     return paras
