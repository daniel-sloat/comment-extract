"""Comments combined"""

from dataclasses import asdict
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
    
    def __len__(self):
        return len(self._comments)

    @cache
    def _remove_replies(self):
        return {
            _id: {"start": comment.start, "end": comment.end}
            for _id, comment in self.comment_bounds.items()
            if _id not in self.ancestors
        }

    def _merge_comment_data(self):
        metadata = {
            _id: data.attrib | {"bubble": data.bubble}
            for _id, data in self.metadata().items()
        }
        return [
            self._remove_replies()[_id] | metadata[_id]# | {"filename": self._doc.file}
            for _id in self._remove_replies()
        ]
        
    # def _merge_comment_data(self):
    #     print(asdict(self.metadata()['0']))
    #     return [
    #         asdict(comment) | asdict(self.metadata()[_id])# | {"filename": self._doc.file}
    #         for _id, comment in self.comment_bounds.items()
    #         if _id not in self.ancestors
    #     ]

    @property
    def _comments(self):
        return [Comment(**comment_data) for comment_data in self._merge_comment_data()]
