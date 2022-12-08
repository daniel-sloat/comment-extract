from functools import cached_property
from docx.ooxml_ns import ns
from docx.comments.comment_data import CommentBounds


class CommentsDocument:
    def __init__(self, document):
        self._doc = document
        self.xml = self._doc.xml["word/document.xml"]
        self.comment_ids = self.xml.xpath(
            "./w:body/w:p/w:commentRangeStart/@w:id|"
            "./w:body/w:commentRangeStart/@w:id"
            **ns,
        )
