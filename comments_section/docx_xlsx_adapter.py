from itertools import groupby
import re
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
    def __init__(self, comment_record, workbook, filename_delimiter, add_columns=False, **attrs):
        self.comment_record = comment_record
        self.workbook = workbook
        self.delimiter = filename_delimiter
        self.add_columns = add_columns

    def sheet_design(self):
        columns = [
            ("Comment Number", "total_count", (5, self.align)),
            ("Filename", "filename.path.name", (15, self.align)),
            ("Document Number", "filename.doc_number", (5, self.align)),
            ("Commenter Code", "filename.commenter_code", (15, self.align)),
            ("Folder", "filename.path.parent.name", (10, self.align)),
            ("Document Comment Number", "doc_comment_count", (5, self.align)),
            ("Author", "comment.author", (15, self.align, self.hidden)),
            ("Initials", "comment.initials", (5, self.align, self.hidden)),
            ("Date", "comment.date", (10, self.align, self.hidden)),
            ("Comment Bubble", "comment.bubble.text", (18, self.text_wrap)),
            ("Referenced Text", "self.combine_runs(comment)", (80, self.text_wrap)),
        ]
        if self.add_columns:
            columns.insert(10, ("Comment Bubble - Category", "comment.bubble.text", (15, self.text_wrap)))
            columns.insert(11, ("Comment Bubble - Alt Commenter Code", "comment.bubble.text", (15, self.text_wrap, self.hidden)))
            columns.insert(13, ("Response", "''", (80, self.text_wrap)))
        return columns

    def header(self):
        return [x for x in self.sheet_design()]

    def data(self):
        total_count = 0
        for comments in self.comment_record:
            filename = FileNameParser(comments._doc.file, self.delimiter)
            doc_comment_count = 0
            for comment in comments:
                q = {}
                total_count += 1
                doc_comment_count += 1
                for col_name, expr, _ in self.sheet_design():
                    q[col_name] = eval(expr)
                if q["Referenced Text"]:
                    yield q

    @staticmethod
    def clean(text):
        text = re.sub(r"\s{2,}", " ", text)  # Replace 2 spaces or more with single
        text = re.sub(r"\s", " ", text)  # Replace any whitespace with regular space
        text = re.sub(r"“|”", '"', text)  # Replace curly quote characters
        return text

    def combine_runs(self, comment):
        runs = []
        for paragraph in comment.paragraphs:
            newline = "\n" if paragraph is not comment.paragraphs[-1] else ""
            for key_format, group in groupby(paragraph.runs, lambda x: x.asdict()):
                text = "".join(run.text for run in group)
                text = self.clean(text)
                if text:
                    runs.extend((self.workbook.add_format(key_format), text))
            if newline:
                runs.append(newline)
        return runs

    def combine_runs(self, comment):
        runs = []
        for paragraph in comment.paragraphs:
            newline = "\n" if paragraph is not comment.paragraphs[-1] else ""
            for key_format, group in groupby(paragraph.runs, lambda x: x.asdict()):
                text_list = []
                group = list(group)
                for run in group:
                    if run == group[0]:
                        run_text = run.text.lstrip()
                        text_list.append(run_text)
                    else:
                        text_list.append(run.text)
                text = "".join(text_list)
                text = self.clean(text)
                if text:
                    runs.extend((self.workbook.add_format(key_format), text))
            if newline:
                runs.append(newline)
        return runs
