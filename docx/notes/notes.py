from docx.ooxml_ns import ns
from docx.notes.note_element import NoteElement


class Notes:
    def __init__(self, document):
        self._doc = document
        self._footnotes_xml = self._doc.xml.get("word/footnotes.xml")
        self._endnotes_xml = self._doc.xml.get("word/endnotes.xml")
        self.notes = {"endnotes": self.endnotes, "footnotes": self.footnotes}

    def __repr__(self):
        return (
            f"Notes("
            f"footnote_count={sum(1 for fn in self.footnotes if int(fn) > 0)},"
            f"endnote_count={sum(1 for en in self.endnotes if int(en) > 0)})"
        )

    def __getitem__(self, key):
        return self.notes[key]

    def __iter__(self):
        return iter(self.notes.items())

    @property
    def endnotes(self):
        if self._endnotes_xml is not None:
            return {
                el.xpath("string(@w:id)", **ns): NoteElement(el, self)
                for el in self._endnotes_xml.xpath("w:endnote", **ns)
            }
        return {}

    @property
    def footnotes(self):
        if self._footnotes_xml is not None:
            return {
                el.xpath("string(@w:id)", **ns): NoteElement(el, self)
                for el in self._footnotes_xml.xpath("w:footnote", **ns)
            }
        return {}
