from .write_xlsx import WriteXLSX
from .filenameparser import FileNameParser


class CommentRecord(list):
    def __repr__(self):
        return f"CommentRecord(file_count={self.__len__()},total_comments={self.total})"

    @property
    def total(self):
        return sum(len(comments) for comments in self)

    def to_excel(self, filename, *args, **kwargs):
        xlsx = WriteXLSX(filename, self, "Comments")
        xlsx.create_workbook()
