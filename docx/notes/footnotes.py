from ..ooxml_ns import ns
from ..elements import Paragraph
from .note_element import NoteElement


class Footnotes:
    def __init__(self, document):
        self._doc = document
        self._footnotes_xml = self._doc.xml.get("word/footnotes.xml")

    def __getitem__(self, key):
        return self.footnotes[key]

    def __iter__(self):
        return iter(self.footnotes)

    @property
    def footnotes(self):
        if self._footnotes_xml is not None:
            return {
                self._footnotes_xml.xpath(
                    "string(w:footnote/@w:id)", **ns
                ): NoteElement(el)
                for el in self._footnotes_xml.find("w:footnote", **ns)
            }
        return {}

    @property
    def paragraphs(self):
        return [Paragraph(el) for el in self.footnotes]
