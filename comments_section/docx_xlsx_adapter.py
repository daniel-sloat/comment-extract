from itertools import groupby
from comments_section.filenameparser import FileNameParser


class XLSXSheetFormats:
    hidden = {"hidden": True}

    @property
    def text_wrap(self):
        return self.workbook.add_format(
            {
                "text_wrap": 1,
                "valign": "top",
                "align": "left",
            }
        )

    @property
    def align(self):
        return self.workbook.add_format(
            {
                "valign": "top",
                "align": "left",
            }
        )

    @property
    def date_format(self):
        return self.workbook.add_format(
            {
                "num_format": "[$-en-US]m/d/yy h:mm AM/PM;@",
                "valign": "top",
                "align": "left",
            }
        )


class DOCX_XLSX_Adapter(XLSXSheetFormats):
    def __init__(self, comment_record, workbook):
        self.comment_record = comment_record
        self.workbook = workbook

    def sheet_design(self, add_columns=None):
        columns = [
            ("Comment Number", "total_count", (5, self.align)),
            ("Filename", "filename.path.name", (15, self.align)),
            ("Document Number", "filename.doc_number", (5, self.align)),
            ("Commenter Code", "filename.commenter_code", (12, self.align)),
            ("Folder", "filename.path.parent.name", (10, self.align)),
            ("Document Comment Number", "doc_comment_count", (5, self.align)),
            ("Author", "comment.author", (15, self.align, self.hidden)),
            ("Initials", "comment.initials", (5, self.align, self.hidden)),
            ("Date", "comment.date", (10, self.align, self.hidden)),
            ("Comment Bubble", "comment.bubble.text", (18, self.text_wrap)),
            ("Referenced Text", "self.combine_runs(comment)", (80, self.text_wrap)),
        ]
        if add_columns:
            columns.insert(10, ("AddCol1", "'Col1'", (18, self.text_wrap)))
            columns.insert(11, ("AddCol2", "'Col2'", (18, self.text_wrap, self.hidden)))
            columns.insert(13, ("AddCol3", "'Col3'", (80, self.text_wrap)))
        return columns

    def header(self):
        return [x for x in self.sheet_design()]

    def data(self):
        total_count = 0
        for comments in self.comment_record:
            filename = FileNameParser(comments._doc.file)
            doc_comment_count = 0
            for comment in comments:
                q = {}
                total_count += 1
                doc_comment_count += 1
                for col_name, expr, _ in self.sheet_design():
                    q[col_name] = eval(expr)
                if q["Referenced Text"]:
                    yield q

    def combine_runs(self, comment):
        runs = []
        for paragraph in comment.paragraphs:
            newline = "\n" if paragraph is not comment.paragraphs[-1] else ""
            for key_format, group in groupby(paragraph.runs, lambda x: x.asdict()):
                text = "".join(run.text for run in group)
                if text:
                    runs.extend((self.workbook.add_format(key_format), text))
            if newline:
                runs.append(newline)
        return runs
