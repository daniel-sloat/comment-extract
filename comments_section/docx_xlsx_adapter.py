from itertools import groupby
from pprint import pprint
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
    def __init__(
        self, comment_record, workbook, filename_delimiter, add_columns=False, **attrs
    ):
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
            columns.insert(
                10,
                (
                    "Comment Bubble - Category",
                    "comment.bubble.text",
                    (15, self.text_wrap),
                ),
            )
            columns.insert(
                11,
                (
                    "Comment Bubble - Alt Commenter Code",
                    "comment.bubble.text",
                    (15, self.text_wrap, self.hidden),
                ),
            )
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

    # def combine_runs(self, comment):
    #     runs = []
    #     for paragraph in comment.paragraphs:
    #         newline = "\n" if paragraph is not comment.paragraphs[-1] else ""
    #         for key_format, group in groupby(paragraph.runs, lambda x: x.asdict()):
    #             text = "".join(run.text for run in group)
    #             text = self.clean(text)
    #             if text:
    #                 runs.extend((self.workbook.add_format(key_format), text))
    #         if newline:
    #             runs.append(newline)
    #     return runs

    class P:
        runs = []
        
        def __iter__(self):
            return iter(self.runs)
    
    def add_notes(self, comment):
        paragraphs = self.P()
        para_runs = []
        footnotes = (run.footnote for paragraph in comment for run in paragraph)
        endnotes = (run.endnote for paragraph in comment for run in paragraph)
        if endnotes or footnotes:
            comment.paragraphs = comment.paragraphs.append()
        return paragraphs

    def combine_runs2(self, comment):
        paragraphs = []
        for paragraph in comment:
            runs = []
            for key_format, run_group in groupby(paragraph.runs, lambda x: x.asdict()):
                run_group_text = "".join(run.text for run in run_group)
                runs.append(
                    (self.workbook.add_format(key_format), self.clean(run_group_text))
                )
            paragraphs.append(runs)
        return paragraphs

    def clean_up(self, comment):
        paragraphs = []
        for paragraph in comment:
            runs = []
            for run in paragraph:
                key_format, text = run
                if run == paragraph[0]:
                    text = text.lstrip()
                if run == paragraph[-1]:
                    text = text.rstrip()
                if not text:
                    continue
                runs.append((key_format, text))
            paragraphs.append(runs)
        return paragraphs

    def para_to_newline(self, comment):
        runs = []
        for paragraph in comment:
            for run in paragraph:
                runs.extend(run)
            if paragraph is not comment[-1]:
                runs.append("\n")
        return runs

    def combine_runs(self, comment):
        # comment = self.add_notes(comment)
        # print(comment)
        comment.paragraphs_with_notes()
        comment = self.combine_runs2(comment)
        comment = self.clean_up(comment)
        comment = self.para_to_newline(comment)
        return comment
