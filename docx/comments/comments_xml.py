from functools import cache
from .comment_data import CommentMetaData
from ..ooxml_ns import ns


class CommentsXML:
    def __init__(self, document):
        self._doc = document

    def __getitem__(self, key):
        return self.metadata[key]

    @cache
    def metadata(self):
        if (xml := self._doc.xml.get("word/comments.xml")) is not None:
            return {
                el.xpath("string(@w:id)", **ns): CommentMetaData(el)
                for el in xml.xpath("w:comment", **ns)
            }
        return {}
