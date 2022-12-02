from comments_section.docx_xlsx_adapter import DOCX_XLSX_Adapter
from comments_section.write_xlsx import WriteComments


class CommentRecord(list):
    def __repr__(self):
        return f"CommentRecord(file_count={len(self)},total_comments={self.total})"

    @property
    def total(self):
        return sum(len(comments) for comments in self)

    # def to_excel(self, output_file, *args, **kwargs):
    #     xlsx = WriteXLSX(output_file, self, "Comments")
    #     xlsx.create_workbook()

    def to_excel(self, output_file, *args, **kwargs):
        adapter = DOCX_XLSX_Adapter(self)
        xlsx = WriteComments(
            filename=output_file,
            header=adapter.header(),
            data=adapter.data(),
        )
        xlsx.create_workbook()
