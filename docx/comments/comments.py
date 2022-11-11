"""Comments combined"""

from .comment_data import Comment
from .comments_doc import CommentsDocument
from .comments_ext import CommentsExt
from .comments_xml import CommentsXML


class Comments(CommentsDocument, CommentsXML, CommentsExt):
    def __init__(self, document):
        super().__init__(document)

    def __repr__(self):
        return f"{len(self.comments)} comments"

    def __getitem__(self, key):
        return self.comments[key]

    def _merge_comment_data(self):
        no_parents = {
            _id: {"start": comment.start, "end": comment.end}
            for _id, comment in self.comment_ranges.items()
            if comment._para_id not in self._parents
        }
        metadata = {
            data.attrib["id"]: data.attrib | {"bubble": data.bubble} for data in self.metadata
        }
        return [no_parents[_id] | metadata[_id] for _id in no_parents]

    @property
    def comments(self):
        return [Comment(**comment_data) for comment_data in self._merge_comment_data()]

    # @property
    # def _ids(self):
    #     return self._doc.xml["word/document.xml"].xpath(
    #         "string(//w:commentRangeStart/@w:id)",
    #         namespaces=ooXMLns,
    #     )

    # @staticmethod
    # def chunks(lst, n):
    #     """Yield successive n-sized chunks from lst."""
    #     for i in range(0, len(lst), n):
    #         yield lst[i : i + n]

    # @property
    # def bounds(self):
    #     return [
    #         CommentBounds(start, end)
    #         for _id in self._ids
    #         for start, end in self.chunks(
    #             self._doc.xml["word/document.xml"].xpath(
    #                 f"//w:commentRangeStart[@w:id='{_id}']|//w:commentRangeEnd[@w:id='{_id}']",
    #                 namespaces=ooXMLns,
    #             ),
    #             2,
    #         )
    #     ]
