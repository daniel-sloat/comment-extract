from ..ooxml_ns import ns
from .comment_data import CommentBounds


class CommentsDocument:
    def __init__(self, document):
        self._doc = document
        self.comment_range_elements = self._doc.xml["word/document.xml"].xpath(
            f".//w:commentRangeStart|.//w:commentRangeEnd", **ns
        )

    @property
    def starts(self):
        return {
            element.xpath("string(@w:id)", **ns): element
            for element in self.comment_range_elements
            if element.xpath("boolean(self::w:commentRangeStart)", **ns)
        }

    @property
    def ends(self):
        return {
            element.xpath("string(@w:id)", **ns): element
            for element in self.comment_range_elements
            if element.xpath("boolean(self::w:commentRangeEnd)", **ns)
        }

    @property
    def comment_ranges(self):
        return {
            _id: CommentBounds(self.starts[_id], self.ends[_id]) for _id in self.starts
        }
