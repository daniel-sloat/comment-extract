from .write_xlsx import WriteXLSX
from .filenameparser import FileNameParser


class CommentRecord(list):
    def to_excel(self, filename, *args, **kwargs):
        xlsx = WriteXLSX(filename, self, "Comments")
        print(xlsx.comments)
