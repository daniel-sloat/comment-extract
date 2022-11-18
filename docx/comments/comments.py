"""Comments combined"""

from functools import cache
from .comment_data import Comment
from .comments_doc import CommentsDocument
from .comments_ext import CommentsExt
from .comments_xml import CommentsXML


class Comments(CommentsDocument, CommentsXML, CommentsExt):
    def __init__(self, document):
        super().__init__(document)

    def __repr__(self):
        return f"Comments(file='{self._doc.file}',count={len(self._comments)})"

    def __getitem__(self, key):
        return self._comments[key]

    @cache
    def _remove_replies(self):
        return {
            _id: {"start": comment.start, "end": comment.end}
            for _id, comment in self.comment_bounds.items()
        }

    def _merge_comment_data(self):
        metadata = {
            _id: data.attrib | {"bubble": data.bubble}
            for _id, data in self.metadata.items()
        }
        return [
            self._remove_replies()[_id] | metadata[_id] | {"filename": self._doc.file}
            for _id in self._remove_replies()
        ]

    @property
    def _comments(self):
        return [Comment(**comment_data) for comment_data in self._merge_comment_data()]
