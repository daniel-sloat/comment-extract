from comments_section.write_xlsx import WriteComments


class CommentRecord(list):
    def __repr__(self):
        return f"CommentRecord(file_count={len(self)},total_comments={self.total})"

    @property
    def total(self):
        return sum(len(comments) for comments in self)

    def to_excel(self, output_file, **attrs):
        xlsx = WriteComments(
            filename=output_file,
            comments=self,
            **attrs,
        )
        xlsx.create_workbook()
