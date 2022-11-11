"""Class Document"""

from functools import cached_property
from zipfile import ZipFile

from lxml import etree

from .comments.comments import Comments
from .filenameparser import FileNameParser
from .styles.styles import Styles


class Document:
    """Opens docx document and creates XML file tree"""
    def __init__(self, filename):
        self.file = FileNameParser(filename)
        self.comments = Comments(self)
        self.styles = Styles(self)

    @cached_property
    def xml(self):
        with ZipFile(self.file.path, "r") as z:
            return {
                filename: etree.fromstring(z.read(filename))
                for filename in z.namelist()
                if ".xml" in filename
            }
