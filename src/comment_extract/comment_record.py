"""Comment record"""

from comment_extract.logger.logger import log_total_record
from comment_extract.write.write_xlsx import WriteComments

from docx_comments.comments.comments import Comments


class CommentRecord:
    def __init__(self, comments: Comments):
        self.comments = comments

    @property
    def total(self):
        return sum(len(comments) for comments in self.comments)

    @log_total_record
    def to_excel(self, output_file: str, **config):
        xlsx = WriteComments(
            filename=output_file,
            comments=self.comments,
            **config,
        )
        xlsx.create_workbook()
