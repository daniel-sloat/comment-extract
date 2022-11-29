from pprint import pprint
from .write_xlsx import WriteXLSX
from .filenameparser import FileNameParser


class CommentRecord(list):
    def __repr__(self):
        return f"CommentRecord(file_count={self.__len__()},total_comments={self.total})"

    @property
    def total(self):
        return sum(len(comments) for comments in self)

    def to_excel(self, output_file, *args, **kwargs):
        xlsx = WriteXLSX(output_file, self, "Comments")
        # pprint(xlsx.prepared_data(), sort_dicts=False)
        # xlsx.prepared_data()
        xlsx.create_workbook()
