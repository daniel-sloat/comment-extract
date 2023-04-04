"""Comments to XLSX adapter"""

from typing import TYPE_CHECKING

from xlsxwriter.workbook import Workbook

from comment_extract.write.xl_comment import XlComment  # pylint: disable=unused-import

if TYPE_CHECKING:
    from comment_extract.comment_record import CommentRecord


class XLSXSheetFormats:
    """Mixin for workbook formats."""

    # pylint: disable=no-member

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


class CommentsAdapter(XLSXSheetFormats):
    """Comments adapter from DOCX to XLSX."""

    def __init__(
        self,
        comment_record: "CommentRecord",
        workbook: Workbook,
        add_columns: bool = False,
        **config,
    ):
        self.comment_record = comment_record
        self.workbook = workbook
        self.add_columns = add_columns
        self.config = config

    def sheet_design(self):
        columns = [
            ("No", "total_count", (5, self.align)),
            ("Filename", "comments._doc.file.name", (17, self.align)),
            ("Folder", "comments._doc.file.parent.name", (8, self.align)),
            ("File Comment Number", "doc_comment_count", (6, self.align)),
            ("Author", "comment.author", (15, self.align, self.hidden)),
            ("Initials", "comment.initials", (6, self.align, self.hidden)),
            ("Date", "comment.date", (15, self.align, self.hidden)),
            ("Comment Bubble", "comment.bubble.text", (18, self.text_wrap)),
            (
                "Referenced Text",
                "XlComment(comment, self.workbook, **self.config).runs",
                (80, self.text_wrap),
            ),
        ]
        if self.add_columns:
            columns.append(("Heading 1", "''", (15, self.text_wrap)))
            columns.append(("Heading 2", "''", (15, self.text_wrap)))
            columns.append(("Response", "''", (80, self.text_wrap)))
        return columns

    def header(self):
        return self.sheet_design()

    def data(self):
        total_count = 0
        for comments in self.comment_record:
            doc_comment_count = 0
            for comment in comments:  # pylint: disable=unused-variable
                comment_data = {}
                total_count += 1
                doc_comment_count += 1
                for col_name, expr, _ in self.sheet_design():
                    comment_data[col_name] = eval(expr)  # pylint: disable=eval-used
                if comment_data["Referenced Text"]:  # Do not include empty comments
                    yield comment_data
