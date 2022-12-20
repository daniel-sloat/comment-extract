"""Class Document"""

from functools import cached_property
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from docx.comments.comments import Comments
from docx.notes.notes import Notes
from docx.styles.styles import Styles
from logger.logger import log_filename


@log_filename
class Document:
    """Opens docx document and creates XML file tree"""

    def __init__(self, filename):
        self.file = Path(filename)
        self.styles = Styles(self)
        self.notes = Notes(self)
        self.comments = Comments(self)

    def __repr__(self):
        return f"Document(file='{self.file}')"

    @cached_property
    def xml(self):
        with ZipFile(self.file, "r") as z:
            return {
                filename: etree.fromstring(z.read(filename))
                for filename in z.namelist()
                if filename
                in (
                    "word/document.xml",
                    "word/styles.xml",
                    "word/comments.xml",
                    "word/commentsExtended.xml",
                    "word/footnotes.xml",
                    "word/endnotes.xml",
                )
            }
