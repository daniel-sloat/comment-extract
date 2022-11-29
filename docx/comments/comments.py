"""Comments combined"""

from docx.comments.comment_data import Comment
from docx.comments.comments_doc import CommentsDocument
from docx.comments.comments_ext import CommentsExt
from docx.comments.comments_xml import CommentsXML

# from logger.logger import log_comments


# @log_comments
class Comments(CommentsDocument, CommentsXML, CommentsExt):
    def __repr__(self):
        return f"Comments(file='{self._doc.file}',count={len(self.comments)})"

    def __iter__(self):
        return iter(self.comments)

    def __len__(self):
        return len(self.comments)

    @property
    def comments(self):
        return [
            Comment(
                **self.comment_bounds[_id],
                **self.metadata()[_id],
                comments=self,
            )
            for _id in self.comment_bounds
            if _id not in self.ancestors
        ]
