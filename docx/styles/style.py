from collections import ChainMap
from functools import cached_property

from docx.elements.attrib import get_attrib
from docx.ooxml_ns import ns


class Style:
    """Representation of <w:style> OOXML document element."""

    def __init__(self, _id, styles):
        self._id = _id
        self._parent = styles
        self.element = self._parent._style_xml.xpath(
            "w:style[@w:styleId=$_id]", _id=self._id, **ns
        )[0]
        self.basedon = self.element.xpath("string(w:basedOn/@w:val)", **ns)

    def __repr__(self):
        return f"Style(_id='{self._id}',type='{self._type}')"

    @property
    def _name(self):
        return self.element.xpath("string(w:name/@w:val)", **ns)

    @property
    def _type(self):
        return self.element.xpath("string(@w:type)", **ns)

    @property
    def _paragraph(self):
        return get_attrib(self.element.xpath("w:pPr/*", **ns))

    @property
    def _run(self):
        return get_attrib(self.element.xpath("w:rPr/*", **ns))

    def _style_inheritance(self):
        based_on_list = []
        based_on = self.basedon
        while based_on:
            based_on_list.append(based_on)
            following_style = self._parent[based_on].basedon
            based_on = following_style
        return based_on_list

    @cached_property
    def paragraph(self):
        props = (
            self._parent[based_on_style]._paragraph
            for based_on_style in self._style_inheritance()
        )
        return ChainMap(self._paragraph, *props, self._parent.doc_default_props_para)

    @cached_property
    def run(self):
        props = (
            self._parent[based_on_style]._run
            for based_on_style in self._style_inheritance()
        )
        return ChainMap(self._run, *props, self._parent.doc_default_props_run)
