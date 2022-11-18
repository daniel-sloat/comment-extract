from .footnotes import Footnotes
from .endnotes import Endnotes


class Notes:
    FOOTNOTEXML = "word/footnotes.xml"
    ENDNOTESXML = "word/endnotes.xml"

    def __init__(self, document):
        self._doc = document
        if self.FOOTNOTEXML in self._doc.xml:
            self.footnotes = Footnotes(self._doc)
        else:
            self.footnotes = []
        if self.ENDNOTESXML in self._doc.xml:
            self.endnotes = Endnotes(self._doc)
        else:
            self.endnotes = []

    def __getitem__(self, key):
        return self.endnotes[key]

    def __iter__(self):
        return iter(self.endnotes)
