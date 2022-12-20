"""Module for basic DOCX element."""
from reprlib import Repr

from lxml.etree import QName

limit = Repr()


class DOCXElement:
    """Basic representation of DOCX Element."""

    def __init__(self, element):
        self.element = element

    def __repr__(self):
        return f"{self.__class__.__name__}(tag='{(self.tag)}')"

    @property
    def tag(self):
        _tag = QName(self.element).localname
        return _tag if _tag != "id" else "_id"

    @property
    def uri(self):
        return QName(self.element).namespace
