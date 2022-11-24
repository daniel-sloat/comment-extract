from pathlib import Path
import xlsxwriter as xl


class WriteXLSX:
    def __init__(self, filename, comment_record, sheetname="Comments", add_columns=False):
        self.filename = filename
        self.comment_record = comment_record
        self.sheetname = sheetname
        self.workbook = xl.Workbook(filename)
        self.worksheet = self.workbook.add_worksheet(self.sheetname)
        self._add_columns = add_columns

    def create_formats(self, run):
        """Encodes format-string to xlsxwriter workbook Format."""
        text, props = run
        run_format = {}
        for p in props:
            match p:
                case "b":
                    run_format["bold"] = True
                case "i":
                    run_format["italic"] = True
                case "u":
                    run_format["underline"] = 1
                case "w":
                    run_format["underline"] = 2
                case "s":
                    run_format["font_strikeout"] = True
                case "z":  # No double-strikeout in Excel so strikeout and red text is used.
                    run_format["font_strikeout"] = True
                    run_format["font_color"] = "#FF0000"  # #FF0000 = red
                case "x":
                    run_format["font_script"] = 1
                case "v":
                    run_format["font_script"] = 2
        format = self.workbook.add_format(run_format)
        run = [format, text]
        return run

    def prepared_data(self):
        p = []
        for comments in self.comment_record:
            for comment in comments:
                for paragraph in comment.paragraphs:
                    for run in paragraph.runs:
                        p.append(run)
                    if paragraph != comment.paragraphs[-1]:
                        p.append(["\n", ""])
        return p

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

    def write_rich_list(self, row: int, col: int, rich_list):
        """A write handler rich formatted lists for xlsxwriter."""
        if len(rich_list) == 2:
            plain_text = rich_list[1]
            return self.worksheet.write_string(row, col, plain_text)
        else:
            return self.worksheet.write_rich_string(row, col, *rich_list)

    def create_workbook(self):
        self.worksheet.add_write_handler(list, self.write_rich_list)

        column_names = ["Comment Number"]  # + dict_keys_as_col_names

        (
            self.set_column_formats()
            .write_header(column_names)
        )

        # for comment_no, comment in enumerate(self.prepared_data()):
        #     row = comment_no + 1
        #     self.worksheet.write(row, 0, row)
        #     for col_num, col_name, data in enumerate(comment.items()):
        #         self.worksheet.write(row, col_num, data[col_name])
        
        # Autofilter and Freeze Panes
        max_row = len(self.prepared_data())
        max_col = len(column_names) - 1
        self.worksheet.autofilter(0, 0, max_row, max_col)
        self.worksheet.freeze_panes(1, 0)

        Path(self.filename).parent.mkdir(exist_ok=True)
        self.workbook.close()
