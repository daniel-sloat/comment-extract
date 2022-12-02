from itertools import groupby
from pathlib import Path
import xlsxwriter as xl


class WriteXLSX:
    def __init__(
        self, filename, comment_record, sheetname="Comments", add_columns=False
    ):
        self.filename = filename
        self.comment_record = comment_record
        self.sheetname = sheetname
        self.workbook = xl.Workbook(filename)
        self.worksheet = self.workbook.add_worksheet(self.sheetname)
        self._add_columns = add_columns

    def create_formats(self, run_format):
        return self.workbook.add_format(run_format)

    # def docx_to_xlsx(self, comment):
    #     for paragraph in comment.paragraphs:
    #         newline = "\n" if paragraph is not comment.paragraphs[-1] else ""
    #         for key_format, group in groupby(
    #             paragraph.runs, lambda x: x.asdict()
    #         ):

    # runs = []
    # for paragraph in comment.paragraphs:
    #     newline = "\n" if paragraph is not comment.paragraphs[-1] else ""
    #     for key_format, group in groupby(
    #         paragraph.runs, lambda x: x.asdict()
    #     ):
    #         text_list = []
    #         for run in group:
    #             text_list.append(run.text)
    #         text = "".join(text_list).strip()
    #         if text:
    #             g = (self.create_formats(key_format), text)
    #             runs.extend(g)
    #         if newline:
    #             runs.append(newline)
    def prepared_data(self):
        p = []
        count = 0
        for comments in self.comment_record:
            path = comments._doc.file
            doc_comment_count = 0
            for comment in comments:
                q = {}

                count += 1
                doc_comment_count += 1

                q["filename"] = path.name
                q["folder"] = path.parent.name
                q["comment_count"] = count
                q["doc_comment_count"] = doc_comment_count
                q["author"] = comment.author
                q["initials"] = comment.initials
                q["date"] = comment.date
                q["bubble"] = comment.bubble.text

                runs = []
                for paragraph in comment.paragraphs:
                    newline = "\n" if paragraph is not comment.paragraphs[-1] else ""
                    for key_format, group in groupby(
                        paragraph.runs, lambda x: x.asdict()
                    ):
                        text = "".join(run.text for run in group)
                        if text:
                            runs.extend((self.create_formats(key_format), text))
                    if newline:
                        runs.append(newline)
                q["runs"] = runs
                if runs:
                    yield q

    def set_column_formats(self):
        """Set column width and cell formats"""
        text_wrap = self.workbook.add_format(
            {"text_wrap": 1, "valign": "top", "align": "left"}
        )
        align = self.workbook.add_format({"valign": "top", "align": "left"})
        date_format = self.workbook.add_format(
            {
                "num_format": "[$-en-US]m/d/yy h:mm AM/PM;@",
                "valign": "top",
                "align": "left",
            }
        )
        hidden = {"hidden": True}

        set_col_formats = [
            (5, align),
            (15, align),
            (5, align),
            (12, align),
            (5, align),
            (18, text_wrap, hidden),
            (5, align, hidden),
            (16, date_format, hidden),
            (18, text_wrap, hidden),
            (18, text_wrap),
            (80, text_wrap),
        ]

        if self._add_columns:
            set_col_formats.insert(10, (18, text_wrap))
            set_col_formats.insert(11, (18, text_wrap, hidden))
            set_col_formats.insert(13, (80, text_wrap))

        for no, col in enumerate(set_col_formats, 0):
            self.worksheet.set_column(no, no, *col)

        return self

    def write_header(self, column_names):
        """Writes header cells with format."""
        header_format = self.workbook.add_format(
            {
                "bold": True,
                "text_wrap": True,
                "valign": "bottom",
                "border": 1,
            }
        )
        for col_num, value in enumerate(column_names):
            self.worksheet.write(0, col_num, value, header_format)
        return self

    @staticmethod
    def write_rich_list(worksheet, row: int, col: int, data):
        """A write handler for rich formatted lists xlsxwriter."""
        if len(data) == 2:
            return worksheet.write_string(row, col, data[1])
        else:
            return worksheet.write_rich_string(row, col, *data)

    def create_workbook(self):
        self.worksheet.add_write_handler(list, self.write_rich_list)

        column_names = ["Comment Number", "Runs"]  # + dict_keys_as_col_names
        self.set_column_formats()
        self.write_header(column_names)

        for comment_no, comment in enumerate(self.prepared_data()):
            row = comment_no + 1
            self.worksheet.write(row, 0, row)
            for col_num, (col_name, data) in enumerate(comment.items(), 1):
                self.worksheet.write(row, col_num, data)

        # Autofilter and Freeze Panes
        max_row = row
        max_col = len(column_names) - 1
        self.worksheet.autofilter(0, 0, max_row, max_col)
        self.worksheet.freeze_panes(1, 0)

        Path(self.filename).parent.mkdir(exist_ok=True)
        self.workbook.close()


class XLSXBase:
    def __init__(self):
        self.workbook = xl.Workbook()


class WriteComments(XLSXBase):
    def __init__(
        self,
        header,
        data,
        filename="output/comments.xlsx",
        sheetname="Comments",
    ):
        super().__init__()
        self.header = header
        self.data = data
        self.filename = filename
        self.sheetname = sheetname
        self.worksheet = self.workbook.add_worksheet(self.sheetname)

    def write_header_and_set_columns(self):
        header_format = self.workbook.add_format(
            {"bold": True, "text_wrap": True, "valign": "bottom", "border": 1}
        )
        for col_num, (col_name, _, formats) in enumerate(self.header):
            self.worksheet.set_column(col_num, col_num, *formats)
            self.worksheet.write(0, col_num, col_name, header_format)

    @staticmethod
    def write_rich_list(worksheet, row: int, col: int, data):
        if len(data) == 2:
            return worksheet.write_string(row, col, data[1])
        else:
            return worksheet.write_rich_string(row, col, *data)

    def create_workbook(self):
        self.worksheet.add_write_handler(list, self.write_rich_list)
        # self.write_header_and_set_columns()

        for row, comment in enumerate(self.data, 1):
            for col_num, (_, data) in enumerate(comment.items()):
                self.worksheet.write(row, col_num, data)

        # Autofilter and Freeze Panes
        max_row = row
        max_col = len(self.header) - 1
        self.worksheet.autofilter(0, 0, max_row, max_col)
        self.worksheet.freeze_panes(1, 0)

        Path(self.filename).parent.mkdir(exist_ok=True)
        self.workbook.filename = self.filename
        self.workbook.close()
