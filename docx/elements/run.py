"""Module for the run <w:r> element."""
from functools import cache

from docx.elements.attrib import get_attrib
from docx.elements.element_base import DOCXElement
from docx.elements.properties import Properties
from docx.ooxml_ns import ns


@cache
class Run(DOCXElement):
    """Representation of run <w:r> element."""

    def __init__(self, element, paragraph):
        super().__init__(element)
        self._parent = paragraph
        self.text = self.element.xpath("string(w:t)", **ns)
        self._props = get_attrib(self.element.xpath("w:rPr/*", **ns))

    def __str__(self):
        return self.text

    @property
    def props(self):
        return Properties(self)

    @props.setter
    def props(self, prop_dict):
        self._props = prop_dict

    @property
    def footnote(self):
        note_id = self.element.xpath("string(w:footnoteReference/@w:id)", **ns)
        return self._parent._doc.notes.footnotes.get(note_id, [])

    @property
    def endnote(self):
        note_id = self.element.xpath("string(w:endnoteReference/@w:id)", **ns)
        return self._parent._doc.notes.endnotes.get(note_id, [])
