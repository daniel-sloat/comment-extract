"""Comments combined"""

from docx.comments.comment_data import Comment
from docx.comments.comments_doc import CommentsDocument
from docx.comments.comments_ext import CommentsExt
from docx.comments.comments_xml import CommentsXML


class Comments(CommentsDocument, CommentsXML, CommentsExt):
    def __init__(self, document):
        super().__init__(document)

    def __repr__(self):
        return f"Comments(file='{self._doc.file}',count={len(self.comments)})"

    def __getitem__(self, key):
        return self.comments[key]

    def __len__(self):
        return len(self.comments)

    @property
    def comments(self):
        return [
            Comment(**self.comment_bounds[_id], **self.metadata()[_id])
            for _id in self.comment_bounds
            if _id not in self.ancestors
        ]
