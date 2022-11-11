from .comment_data import CommentMetaData
from ..ooxml_ns import *


class CommentsXML:
    def __init__(self, document):
        self._doc = document

    @property
    def _comments_xml(self):
        return self._doc.xml.get("word/comments.xml")

    @property
    def metadata(self):
        if self._comments_xml is not None:
            return [
                CommentMetaData(element)
                for element in self._comments_xml.xpath("w:comment", **ns)
            ]
        return []
