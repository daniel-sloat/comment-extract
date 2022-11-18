from ..ooxml_ns import ns
from ..elements import Paragraph
from .note_element import NoteElement


class Endnotes:
    def __init__(self, document):
        self._doc = document
        self._xml = self._doc.xml.get("word/endnotes.xml")

    def __getitem__(self, key):
        return self.endnotes[key]

    def __iter__(self):
        return iter(self.endnotes)

    @property
    def endnotes(self):
        if self._xml is not None:
            return {
                self._xml.xpath("string(w:footnote/@w:id)", **ns): NoteElement(el)
                for el in self._xml.find("w:endnote", **ns)
            }
        return {}

    @property
    def paragraphs(self):
        return [Paragraph(el) for el in self.endnotes]
