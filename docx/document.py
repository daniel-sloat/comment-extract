"""Class Document"""

from functools import cached_property
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from .comments.comments import Comments
from .styles.styles import Styles
from .notes.notes import Notes


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
                if ".xml" in filename
            }
