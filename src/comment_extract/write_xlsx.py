from pathlib import Path

import xlsxwriter as xl

from comment_extract.adapter import CommentsAdapter


class WriteComments:
    def __init__(
        self, comments, filename="output/comments.xlsx", sheetname="Comments", **config
    ):
        self.filename = filename
        self.workbook = xl.Workbook(self.filename)
        self.comments = CommentsAdapter(comments, self.workbook, **config)
        self.sheetname = sheetname
        self.worksheet = self.workbook.add_worksheet(self.sheetname)
        self.worksheet.add_write_handler(list, self.write_rich_list)

    @staticmethod
    def write_rich_list(worksheet, row, col, data):
        if len(data) == 2:
            _, text = data
            return worksheet.write(row, col, text)
        else:
            return worksheet.write_rich_string(row, col, *data)

    def create_workbook(self):
        header_format = self.workbook.add_format(
            {"bold": True, "text_wrap": True, "valign": "bottom", "border": 1}
        )

        for col_num, (col_name, _, formats) in enumerate(self.comments.header()):
            self.worksheet.set_column(col_num, col_num, *formats)
            self.worksheet.write(0, col_num, col_name, header_format)

        for row_num, comment in enumerate(self.comments.data(), 1):
            for col_num, (_, data) in enumerate(comment.items()):
                self.worksheet.write(row_num, col_num, data)

        self.worksheet.autofilter(0, 0, row_num, col_num)
        self.worksheet.freeze_panes(1, 0)

        Path(self.filename).parent.mkdir(exist_ok=True)
        self.workbook.close()
