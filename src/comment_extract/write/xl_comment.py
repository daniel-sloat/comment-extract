"""Prepare comment for Excel"""

import re
from itertools import groupby

from docx_comments.comments.comment import Comment
from docx_comments.docx import Document
from docx_comments.elements.element_maker import paragraph_maker
from docx_comments.elements.paragraph import Paragraph
from xlsxwriter.workbook import Workbook

from comment_extract.write.xl_run import XlRun


def group_runs(runs):
    return [
        (props, "".join(run[1] for run in run_group))
        for props, run_group in groupby(runs, lambda x: x[0])
    ]


class XlComment:
    """Adapter between DOCX Comment object and XLSXWriter"""

    def __init__(self, comment: Comment, workbook: Workbook, **config):
        self._docxcomment = comment
        self._doc: Document = self._docxcomment._parent._doc
        self.workbook = workbook
        self.config = config

    @property
    def runs(self):
        formatted_runs = []
        paragraphs = self._docxcomment.paragraphs
        paragraphs.extend(self.append_notes())
        for paragraph in paragraphs:
            xl_runs = [(XlRun(run).props, run.text) for run in paragraph]
            if paragraph != paragraphs[-1]:
                xl_runs.append(({}, "\n"))
            xl_runs = group_runs(xl_runs)
            for props, text in self._clean_paragraph_text(xl_runs):
                formatted_runs.extend((self.workbook.add_format(props), text))
        return formatted_runs

    @staticmethod
    def _clean_paragraph_text(runs: list[tuple[dict, str]]):
        for run in runs:
            props, text = run
            # Attempt to remove leading and trailing whitespace at the beginngin and
            # end of a paragraph (e.g., it's possible for second run to contain leading
            # whitespace that isn't removed.)
            if run == runs[0]:
                text = text.lstrip()
            if run == runs[-1]:
                text = text.rstrip(" ")  # Specific, to not remove '\n'
            if not text:  # XLSXWriter cannot write an empty run
                continue

            text = re.sub(r"[^\S\r\n]+", " ", text)  # Replace whitespace with reg space
            text = re.sub(r"“|”", '"', text)  # Replace curly quote characters

            yield props, text

    def note_handler(self, note_data, note_num):
        superscript = {"vertAlign": {"val": "superscript"}}
        for para_no, paragraph in enumerate(note_data.paragraphs):
            runs = []
            if para_no == 0:
                runs.append((superscript, note_num))
            for run in paragraph:
                if run.text:
                    runs.append((run.props.chain, run.text))
            para_element = paragraph_maker(*runs)
            yield Paragraph(para_element, self._doc)

    def append_notes(self):
        superscript = {"vertAlign": {"val": "superscript"}}
        count = 0
        for paragraph in self._docxcomment:
            for run in paragraph:
                if run.footnote:
                    count += 1
                    run.text = str(count)
                    run.props = superscript
                    for para in self.note_handler(run.footnote, str(count)):
                        yield para
                elif run.endnote:
                    count += 1
                    run.text = str(count)
                    run.props = superscript
                    for para in self.note_handler(run.endnote, str(count)):
                        yield para
