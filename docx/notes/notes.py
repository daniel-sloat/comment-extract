from lxml.etree import Element

from docx.notes.note import EndNote, FootNote
from docx.ooxml_ns import ns


class Notes:
    """Representation of footnotes and endnotes in OOXML."""

    def __init__(self, document):
        self._doc = document
        self._footnotes_xml = self._doc.xml.get("word/footnotes.xml", Element("root"))
        self._endnotes_xml = self._doc.xml.get("word/endnotes.xml", Element("root"))
        self.footnote_ids = self._footnotes_xml.xpath("w:footnote/@w:id", **ns)
        self.endnote_ids = self._endnotes_xml.xpath("w:endnote/@w:id", **ns)
        self.footnotes = {_id: FootNote(_id, self) for _id in self.footnote_ids}
        self.endnotes = {_id: EndNote(_id, self) for _id in self.endnote_ids}
        self.notes = {"endnotes": self.endnotes, "footnotes": self.footnotes}

    def __repr__(self):
        return (
            f"Notes("
            f"footnotes={sum(1 for _id in self.footnote_ids if int(_id) > 0)},"
            f"endnotes={sum(1 for _id in self.endnote_ids if int(_id) > 0)})"
        )

    def __getitem__(self, key):
        return self.notes[key]

    def __iter__(self):
        return iter(self.notes.items())

    def __len__(self):
        return len(self.footnote_ids) + len(self.endnote_ids)
