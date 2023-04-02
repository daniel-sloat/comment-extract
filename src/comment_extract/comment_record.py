"""Comment record"""

from comment_extract.logger.logger import log_total_record
from comment_extract.write.write_xlsx import WriteComments


class CommentRecord(list):
    """The comment record, subclassed from list."""

    def __repr__(self):
        return f"CommentRecord(file_count={len(self)},total_comments={self.total})"

    @property
    def total(self):
        return sum(len(comments) for comments in self)

    @log_total_record
    def to_excel(self, output_file, **config):
        xlsx = WriteComments(
            filename=output_file,
            comments=self,
            **config,
        )
        xlsx.create_workbook()
