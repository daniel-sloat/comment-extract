from comments_section.write_xlsx import WriteComments
from logger.logger import log_total_record


class CommentRecord(list):
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
