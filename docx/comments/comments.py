"""Comments combined"""

from docx.comments.comment import Comment
from docx.ooxml_ns import ns
from logger.logger import log_comments


@log_comments
class Comments:
    """Comments contained within document. Only top-level comments are included.
    Replies of comments are not."""

    def __init__(self, document):
        self._doc = document
        self._document_root = self._doc.xml["word/document.xml"]
        self._comment_metadata_root = self._doc.xml.get("word/comments.xml")
        self._comment_ext_root = self._doc.xml.get("word/commentsExtended.xml")
        self.comment_ids = self._document_root.xpath(
            "./w:body//w:commentRangeStart/@w:id",
            **ns,
        )
        self._all_comments = [Comment(_id, self) for _id in self.comment_ids]
        self.comments = [
            comment for comment in self._all_comments if not comment._is_reply
        ]

    def __repr__(self):
        return f"Comments(file='{self._doc.file}',count={len(self.comments)})"

    def __getitem__(self, key):
        return self.comments[key]

    def __iter__(self):
        return iter(self.comments)

    def __len__(self):
        return len(self.comments)
