"""Comment data containers"""

from lxml import etree
from dataclasses import dataclass
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
    
# @dataclass
# class CommentMetaData:
#     element: etree._Element
    
#     def __post_init__(self):
#         self._id = self.element.xpath("string(@w:id)", **ns)
#         self.bubble = Bubble(self.element)
#         self.para_id = self.element.xpath("string(w:p[last()]/@w14:paraId)", **ns)


@dataclass
class CommentBounds:
    start: etree._Element
    end: etree._Element

    def __post_init__(self):
        self._id = self.start.xpath("string(@w:id)", **ns)

    # @property
    # def last_para_id(self):
    #     return self.end.xpath(
    #         "string((parent::w:p|preceding-sibling::w:p[1])/@w14:paraId)", **ns
    #     )


class Comment:
    def __init__(
        self,
        # filename,
        *,
        start,
        end,
        _id,
        bubble,
        author="",
        date="",
        initials="",
        **kwargs,
    ):
        # self.filename = filename
        self._start = start
        self._end = end
        self._id = _id
        self.author = author
        self.date = date
        self.initials = initials
        self.bubble = bubble

    def __repr__(self):
        return f"Comment(_id='{self._id}',text='{self.text}')"

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
